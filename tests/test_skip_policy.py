"""당월/28일 정책 전환 및 시스템 관리 저장 검증."""

import contextlib
import io
from pathlib import Path
import tempfile
import unittest

import dangjik_basic1 as dy
from app import app


class SkipPolicyTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.original_db = dy.DB_NAME
        dy.DB_NAME = str(Path(self.tmp.name) / 'policy.db')
        dy.init_db()
        self.emp = dy.register_employee('00000001', '직원', '부서', dy.GROUP_SUK)
        self.client = app.test_client()

    def tearDown(self) -> None:
        dy.DB_NAME = self.original_db
        self.tmp.cleanup()

    def history(self, date: str) -> None:
        with dy.get_conn() as conn:
            conn.execute('INSERT INTO duty_history(duty_date,duty_type,emp_id) VALUES (?,?,?)',
                         (date, '숙직', self.emp))

    def blocked(self, date: str) -> bool:
        with dy.get_conn() as conn:
            return dy._has_recent_regular_duty(conn.cursor(), self.emp, date)

    def test_default_on_and_legacy_database_default(self) -> None:
        self.assertTrue(dy.get_skip_current_month_only())
        html = self.client.get('/admin').get_data(as_text=True)
        self.assertIn('name="skip_current_month_only" checked', html)
        with dy.get_conn() as conn:
            conn.execute('DROP TABLE system_settings')
        self.assertTrue(dy.get_skip_current_month_only())
        dy.init_db()
        self.assertTrue(dy.get_skip_current_month_only())

    def test_cross_month_and_same_month_29_day_cases(self) -> None:
        self.history('2026-04-27')
        self.assertFalse(self.blocked('2026-05-01'))
        dy.set_skip_current_month_only(False)
        self.assertTrue(self.blocked('2026-05-01'))
        with dy.get_conn() as conn:
            conn.execute('DELETE FROM duty_history')
        self.history('2026-01-01')
        self.assertFalse(self.blocked('2026-01-30'))
        dy.set_skip_current_month_only(True)
        self.assertTrue(self.blocked('2026-01-30'))  # PDF 시뮬레이션의 AD 사례

    def test_policy_save_is_persistent_and_does_not_modify_assignments(self) -> None:
        self.history('2026-04-27')
        with dy.get_conn() as conn:
            before = conn.execute('SELECT * FROM duty_history').fetchall()
        response = self.client.post('/admin/skip-policy', data={}, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertFalse(dy.get_skip_current_month_only())
        self.assertNotIn('name="skip_current_month_only" checked', response.get_data(as_text=True))
        dy.init_db()  # 재시작 시 기존 설정 보존
        self.assertFalse(dy.get_skip_current_month_only())
        with contextlib.redirect_stdout(io.StringIO()):
            dy.delete_month('2026-03')
        self.assertFalse(dy.get_skip_current_month_only())
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT * FROM duty_history').fetchall(), before)
        self.client.post('/admin/skip-policy', data={'skip_current_month_only': 'on'})
        self.assertTrue(dy.get_skip_current_month_only())

    def test_penalty_exception_and_current_regular_duty_in_both_modes(self) -> None:
        self.history('2026-05-01')
        dy.insert_manual_override('2026-05-01', dy.GROUP_SUK, '00000001', '대근')
        for enabled in (True, False):
            dy.set_skip_current_month_only(enabled)
            self.assertFalse(self.blocked('2026-05-20'))
            self.assertTrue(self.blocked('2026-05-01'))
        self.history('2026-05-10')
        for enabled in (True, False):
            dy.set_skip_current_month_only(enabled)
            self.assertTrue(self.blocked('2026-05-20'))

    def test_reset_restores_default_checked(self) -> None:
        dy.set_skip_current_month_only(False)
        with contextlib.redirect_stdout(io.StringIO()):
            dy.reset_database(backup=False)
        self.assertTrue(dy.get_skip_current_month_only())


if __name__ == '__main__':
    unittest.main()
