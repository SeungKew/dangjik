"""직원 전입·전출 테스트. 모든 변경은 임시 DB에서만 수행한다."""

import contextlib
import datetime
import io
import json
import sqlite3
import tempfile
import unittest
from pathlib import Path

import dangjik_basic1 as dy
from app import app


class EmployeeChangesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.original_db = dy.DB_NAME
        dy.DB_NAME = str(Path(self.tmp.name) / 'test.db')
        dy.init_db()
        self.client = app.test_client()

    def tearDown(self) -> None:
        dy.DB_NAME = self.original_db
        self.tmp.cleanup()

    def add(self, number: int, group: str = dy.GROUP_SUK) -> int:
        return dy.register_employee(f'{number:08d}', f'직원{number}', '테스트부서', group)

    def test_register_appends_and_preserves_leading_zero(self) -> None:
        self.add(1)
        dy.transfer_out_employee(self.add(2))
        third = self.add(3)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT emp_no,group_idx,is_active FROM employees WHERE id=?',
                                          (third,)).fetchone(), ('00000003', 3, 1))
        self.add(4, dy.GROUP_SEN)
        self.assertEqual(dy.list_employees(dy.GROUP_SEN)[0][2], 1)

    def test_duplicate_and_invalid_registration_do_not_overwrite(self) -> None:
        self.add(1)
        cases = [('00000001', '다른이름', '다른부서', dy.GROUP_SUK),
                 ('00000000', '이름', '부서', dy.GROUP_SUK),
                 ('1234abcd', '이름', '부서', dy.GROUP_SUK),
                 ('123', '이름', '부서', dy.GROUP_SUK),
                 ('00000002', '', '부서', dy.GROUP_SUK),
                 ('00000002', '이름', ' ', dy.GROUP_SUK),
                 ('00000002', '이름', '부서', '잘못된그룹')]
        for args in cases:
            with self.subTest(args=args), self.assertRaises(ValueError):
                dy.register_employee(*args)
        rows = dy.list_employees()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][3], '직원1')

    def test_transfer_preserves_history_manual_queue_and_pointer(self) -> None:
        emp = self.add(1)
        with dy.get_conn() as conn:
            conn.execute('INSERT INTO duty_history(duty_date,duty_type,emp_id) VALUES (?,?,?)',
                         ('2026-02-01', '숙직', emp))
            conn.execute('INSERT INTO system_state VALUES (?,?)', (dy.PTR_FWD, emp))
            conn.execute('INSERT INTO skip_queue(emp_id,skipped_date,return_year_month) VALUES (?,?,?)',
                         (emp, '2026-02-02', '2026-04'))
        dy.insert_manual_override('2026-02-01', dy.GROUP_SUK, '00000001', '대근')
        self.assertTrue(dy.transfer_out_employee(emp))
        self.assertTrue(dy.transfer_out_employee(emp))
        self.assertFalse(dy.transfer_out_employee(999))
        self.assertEqual(dy.list_employees()[0][6], 0)
        self.assertEqual(dy.get_assignments('2026-02')[0][2], '직원1')
        with dy.get_conn() as conn:
            for table in ('duty_history', 'manual_overrides', 'skip_queue', 'system_state'):
                self.assertEqual(conn.execute(f'SELECT COUNT(*) FROM {table}').fetchone()[0], 1)

    def test_departed_pointer_continues_in_both_directions(self) -> None:
        first, second, third = self.add(1), self.add(2), self.add(3)
        dy.transfer_out_employee(second)
        for key, reverse, expected in ((dy.PTR_FWD, False, third), (dy.PTR_REV, True, first)):
            with self.subTest(reverse=reverse), dy.get_conn() as conn:
                conn.execute('REPLACE INTO system_state VALUES (?,?)', (key, second))
                selected, _ = dy.get_next_candidate(conn.cursor(), '2026-03-01', dy.GROUP_SUK, key, reverse)
                self.assertEqual(selected, expected)

    def test_no_active_candidates_does_not_crash(self) -> None:
        for excluded in (False, True):
            with self.subTest(excluded=excluded):
                if excluded:
                    dy.transfer_out_employee(self.add(1))
                with contextlib.redirect_stdout(io.StringIO()):
                    response = self.client.post('/generate', data={'month': '2026-03'}, follow_redirects=True)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(dy.get_assignments('2026-03'), [])

    def test_web_registration_transfer_and_reactivation(self) -> None:
        page = self.client.get('/employees').get_data(as_text=True)
        self.assertIn(dy.GROUP_SUK, page)
        form = dict(emp_no='00000001', name='전입직원', dept='인재경영처', duty_group=dy.GROUP_SUK)
        response = self.client.post('/employees/register', data=form, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn('전입 직원을 등록했습니다', response.get_data(as_text=True))
        emp = dy.list_employees()[0][0]
        duplicate = self.client.post('/employees/register', data=form, follow_redirects=True)
        self.assertIn('이미 등록된 사번', duplicate.get_data(as_text=True))
        response = self.client.post(f'/employees/{emp}/transfer-out', follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(dy.list_employees(status='active'), [])
        self.assertEqual(self.client.get('/employees?status=inactive&q=전입직원').status_code, 200)
        response = self.client.post(f'/employees/{emp}/update', data=dict(
            group_idx='1', name='전입직원', dept='인재경영처', duty_group=dy.GROUP_SUK,
            is_active_sel='1'), follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(dy.list_employees()[0][6], 1)

    def test_existing_v2_march_regression(self) -> None:
        source_path = Path(__file__).resolve().parents[1] / 'duty_system_v2.db'
        baseline_path = Path(__file__).parent / 'fixtures/v2_march_generation_before_employee_changes.json'
        expected = [tuple(row) for row in json.loads(baseline_path.read_text())]
        with sqlite3.connect(source_path.as_uri() + '?mode=ro', uri=True) as source:
            with dy.get_conn() as target:
                source.backup(target)
        dy.init_db()
        dy.set_skip_current_month_only(False)  # 변경 전 생성 fixture는 28일 정책 기준
        with contextlib.redirect_stdout(io.StringIO()):
            dy.regenerate_month('2026-03-01')
        with dy.get_conn() as conn:
            generated = conn.execute("SELECT duty_date,duty_type,emp_id FROM duty_history "
                                     "WHERE duty_date LIKE '2026-03%' ORDER BY 1,2,3").fetchall()
        # Pre-change generation baseline; stored v2 history is not a generation oracle.
        self.assertEqual(len(expected), 51)
        self.assertEqual(generated, expected)

    def test_first_grade_manual_day_duty_replaces_senior_slot(self) -> None:
        self.assertIn(dy.GROUP_FIRST, self.client.get('/employees').get_data(as_text=True))
        response = self.client.post('/employees/register', data=dict(
            emp_no='96100230', name='이재수', dept='테스트부서', duty_group=dy.GROUP_FIRST),
            follow_redirects=True)
        self.assertIn('전입 직원을 등록했습니다', response.get_data(as_text=True))
        self.assertIn('96100230', self.client.get('/manual').get_data(as_text=True))
        self.add(1, dy.GROUP_SEN)
        self.add(2, dy.GROUP_JUN)
        self.add(3, dy.GROUP_SUK)
        self.client.post('/manual', data=dict(target_date='2026-03-07', emp_no='96100230', reason='벌당직'))
        with dy.get_conn() as conn:
            plan = dy.assign_day(conn.cursor(), datetime.date(2026, 3, 7))
            self.assertEqual([(p[0], p[1]) for p in plan],
                             [('일직', '96100230'), ('일직', '00000002'), ('숙직', '00000003')])
            self.assertIsNone(conn.execute('SELECT emp_id FROM system_state WHERE pointer_key=?',
                                           (dy.PTR_SEN,)).fetchone())
            next_plan = dy.assign_day(conn.cursor(), datetime.date(2026, 3, 8))
            self.assertEqual(next_plan[0][1], '00000001')
            self.assertFalse(any(p[1] == '96100230' for p in next_plan))
            self.assertEqual(conn.execute('SELECT COUNT(*) FROM duty_history h JOIN employees e ON e.id=h.emp_id '
                                          'WHERE e.emp_no=?', ('96100230',)).fetchone()[0], 1)


if __name__ == '__main__':
    unittest.main()
