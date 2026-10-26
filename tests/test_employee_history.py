"""직원 변경과 이력 저장의 원자성 및 웹 조회 검증."""

import contextlib
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import openpyxl

import dangjik_basic1 as dy
from app import app


class EmployeeHistoryTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.original_db = dy.DB_NAME
        dy.DB_NAME = str(Path(self.tmp.name) / 'history.db')
        dy.init_db()
        self.client = app.test_client()

    def tearDown(self) -> None:
        dy.DB_NAME = self.original_db
        self.tmp.cleanup()

    def update(self, emp: int, active: int, reason: str = '', name: str = '직원') -> None:
        self.assertTrue(dy.update_employee(emp, dy.GROUP_SUK, 1, name, '부서', active,
                                           reason, '2026-03-01'))

    def test_lifecycle_before_after_reason_and_noop(self) -> None:
        emp = dy.register_employee('00000001', '직원', '부서', dy.GROUP_SUK, '본사 전입', '2026-02-01')
        dy.transfer_out_employee(emp, '지사 전출', '2026-02-20')
        self.update(emp, 1, '본사 재전입')
        self.update(emp, 0, '부서 업무로 제외')
        self.update(emp, 1, '제외 해제')
        self.update(emp, 1, '성명 정정', '새이름')
        self.update(emp, 1, '중복 저장', '새이름')
        rows, total = dy.list_employee_changes(emp_id=emp)
        self.assertEqual(total, 6)
        self.assertEqual([r['change_type'] for r in reversed(rows)],
                         ['전입', '전출', '재전입', '제외', '활성복구', '정보변경'])
        self.assertEqual(rows[0]['reason'], '성명 정정')
        self.assertEqual(json.loads(rows[0]['before_state'])['name'], '직원')
        self.assertEqual(json.loads(rows[0]['after_state'])['name'], '새이름')
        self.assertEqual(rows[-1]['effective_date'], '2026-02-01')
        self.assertTrue(rows[-1]['recorded_at'].endswith('+09:00'))
        self.assertEqual(rows[-1]['name'], '직원')  # 과거 이름 보존

    def test_failed_history_write_rolls_back_employee_change(self) -> None:
        emp = dy.register_employee('00000001', '직원', '부서', dy.GROUP_SUK)
        with patch.object(dy, '_record_employee_change', side_effect=RuntimeError('기록 실패')):
            with self.assertRaises(RuntimeError):
                self.update(emp, 0, '제외')
            with self.assertRaises(RuntimeError):
                dy.register_employee('00000002', '직원2', '부서', dy.GROUP_SUK)
            with self.assertRaises(RuntimeError):
                dy.transfer_out_employee(emp)
        self.assertEqual(dy.list_employees()[0][6], 1)
        self.assertEqual(len(dy.list_employees()), 1)
        self.assertEqual(dy.list_employee_changes()[1], 1)

    def test_excluded_employee_can_transfer_without_duplicate_events(self) -> None:
        emp = dy.register_employee('00000001', '직원', '부서', dy.GROUP_SUK)
        self.update(emp, 0, '일시 제외')
        dy.transfer_out_employee(emp, '전출 확정')
        dy.transfer_out_employee(emp, '중복 클릭')
        rows, total = dy.list_employee_changes(emp_id=emp)
        self.assertEqual(total, 3)
        self.assertEqual(rows[0]['change_type'], '전출')
        self.assertEqual(rows[0]['reason'], '전출 확정')

    def test_invalid_date_is_rejected_without_changes(self) -> None:
        emp = dy.register_employee('00000001', '직원', '부서', dy.GROUP_SUK)
        response = self.client.post(f'/employees/{emp}/transfer-out',
                                    data={'effective_date': '2026-02-30', 'reason': '전출'}, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn('변동일은', response.get_data(as_text=True))
        self.assertEqual(dy.list_employees()[0][6], 1)
        self.assertEqual(dy.list_employee_changes()[1], 1)

    def test_history_page_filters_pagination_and_month_deletion(self) -> None:
        emp = dy.register_employee('00000001', '<직원>', '부서', dy.GROUP_SUK, '<사유>', '2026-02-01')
        self.update(emp, 0, '부서 사정', '<직원>')
        self.update(emp, 1, '복구', '<직원>')
        first, total = dy.list_employee_changes(emp_id=emp, page_size=2)
        second, _ = dy.list_employee_changes(emp_id=emp, page_size=2, page=2)
        self.assertEqual((len(first), len(second), total), (2, 1, 3))
        self.assertNotEqual(first[0]['id'], second[0]['id'])
        response = self.client.get(f'/employees/history?emp_id={emp}&type=제외')
        self.assertEqual(response.status_code, 200)
        self.assertIn('부서 사정', response.get_data(as_text=True))
        self.assertIn('&lt;직원&gt;', response.get_data(as_text=True))
        self.assertEqual(self.client.get('/employees/history?page=abc').status_code, 200)
        with contextlib.redirect_stdout(io.StringIO()):
            dy.delete_month('2026-03')
        self.assertEqual(dy.list_employee_changes(emp_id=emp)[1], 3)

    def test_legacy_schema_does_not_invent_history(self) -> None:
        emp = dy.register_employee('00000001', '직원', '부서', dy.GROUP_SUK)
        with dy.get_conn() as conn:
            conn.execute('DROP TABLE employee_change_history')
        dy.init_db()
        self.assertEqual(dy.list_employee_changes()[1], 0)
        self.update(emp, 0, '부서 제외')
        rows, total = dy.list_employee_changes()
        self.assertEqual(total, 1)
        self.assertEqual(json.loads(rows[0]['before_state'])['is_active'], 1)

    def test_excel_upload_logs_final_change_once(self) -> None:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = '직원명부(마스터)'
        sheet.append(['당직그룹', '해당그룹내순번', '성명', '사번', '부서', '상태(활성/제외)'])
        sheet.append([dy.GROUP_SUK, 1, '직원', '00000001', '부서', '활성'])
        wb.create_sheet('초기당직이력(2월)').append(['당직수행일', '당직분류', '사번'])
        path = Path(self.tmp.name) / 'input.xlsx'
        wb.save(path)
        with contextlib.redirect_stdout(io.StringIO()):
            dy.load_initial_data(str(path))
            dy.load_initial_data(str(path))
        self.assertEqual(dy.list_employee_changes()[1], 1)
        sheet['F2'] = '제외'
        wb.save(path)
        with contextlib.redirect_stdout(io.StringIO()):
            dy.load_initial_data(str(path))
        wb.close()
        rows, total = dy.list_employee_changes()
        self.assertEqual(total, 2)
        self.assertEqual(rows[0]['change_type'], '엑셀변경')
        self.assertEqual(rows[0]['source'], '엑셀 업로드')
        self.assertIn('상태: 활성 → 제외', rows[0]['changes'])

    def test_web_reason_and_date_are_saved_for_status_and_transfer(self) -> None:
        self.client.post('/employees/register', data=dict(
            emp_no='00000001', name='직원', dept='부서', duty_group=dy.GROUP_SUK,
            effective_date='2026-02-01', reason='전입 발령'))
        emp = dy.list_employees()[0][0]
        self.client.post(f'/employees/{emp}/update', data=dict(
            group_idx='1', name='직원', dept='부서', duty_group=dy.GROUP_SUK, is_active_sel='0',
            effective_date='2026-03-01', reason='부서 특성으로 제외'))
        self.client.post(f'/employees/{emp}/transfer-out', data=dict(
            effective_date='2026-04-01', reason='사업소 전출'))
        rows, total = dy.list_employee_changes(emp_id=emp)
        self.assertEqual(total, 3)
        self.assertEqual([(r['change_type'], r['effective_date'], r['reason']) for r in reversed(rows)],
                         [('전입', '2026-02-01', '전입 발령'), ('제외', '2026-03-01', '부서 특성으로 제외'),
                          ('전출', '2026-04-01', '사업소 전출')])
        page = self.client.get('/employees/history?q=00000001').get_data(as_text=True)
        self.assertIn('부서 특성으로 제외', page)
        self.assertIn('사업소 전출', page)


if __name__ == '__main__':
    unittest.main()
