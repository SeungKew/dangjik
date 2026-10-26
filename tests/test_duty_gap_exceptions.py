"""제도 안내 PDF의 벌당직·대근 예외와 일반 28일 제한 검증."""

import datetime
from pathlib import Path
import tempfile
import unittest

import dangjik_basic1 as dy


class DutyGapExceptionsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.original_db = dy.DB_NAME
        dy.DB_NAME = str(Path(self.tmp.name) / 'gap.db')
        dy.init_db()
        dy.set_skip_current_month_only(False)  # 기존 28일 정책의 회귀 검증
        self.first = dy.register_employee('00000001', '첫직원', '부서', dy.GROUP_SEN)
        self.second = dy.register_employee('00000002', '다음직원', '부서', dy.GROUP_SEN)

    def tearDown(self) -> None:
        dy.DB_NAME = self.original_db
        self.tmp.cleanup()

    def history(self, date: str, reason: str | None = None) -> None:
        with dy.get_conn() as conn:
            conn.execute('INSERT INTO duty_history(duty_date,duty_type,emp_id) VALUES (?,?,?)',
                         (date, '일직', self.first))
        if reason is not None:
            dy.insert_manual_override(date, dy.GROUP_SEN, '00000001', reason)

    def choose(self, date: str) -> int | None:
        with dy.get_conn() as conn:
            selected, _ = dy.get_next_candidate(conn.cursor(), date, dy.GROUP_SEN, dy.PTR_SEN)
        return selected

    def test_penalty_and_substitution_do_not_displace_regular_turn(self) -> None:
        for reason in ('벌당직', '대근', '보안점검 위규 벌당직 반영'):
            with self.subTest(reason=reason):
                with dy.get_conn() as conn:
                    conn.execute('DELETE FROM duty_history')
                    conn.execute('DELETE FROM manual_overrides')
                    conn.execute('DELETE FROM system_state')
                    conn.execute('DELETE FROM skip_queue')
                self.history('2026-03-01', reason)
                self.assertEqual(self.choose('2026-03-29'), self.first)
                with dy.get_conn() as conn:
                    self.assertEqual(conn.execute('SELECT COUNT(*) FROM duty_history').fetchone()[0], 1)
                    self.assertEqual(conn.execute('SELECT COUNT(*) FROM skip_queue').fetchone()[0], 0)

    def test_ordinary_duty_still_skips_at_28_days(self) -> None:
        self.history('2026-03-01')
        self.assertEqual(self.choose('2026-03-29'), self.second)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT emp_id FROM skip_queue').fetchone()[0], self.first)

    def test_regular_duty_before_recent_penalty_still_blocks(self) -> None:
        self.history('2026-03-01')
        self.history('2026-03-20', '벌당직')
        self.assertEqual(self.choose('2026-03-29'), self.second)

    def test_ordinary_manual_assignment_is_not_a_penalty(self) -> None:
        self.history('2026-03-01', '일반 지정')
        self.assertEqual(self.choose('2026-03-29'), self.second)

    def test_29_days_and_future_history(self) -> None:
        self.history('2026-03-01')
        self.history('2026-04-20')
        self.assertEqual(self.choose('2026-03-30'), self.first)

    def test_same_day_history_is_not_exempted(self) -> None:
        self.history('2026-03-29', '대근')
        # Check another regular date selection after the prior exception was recorded.
        with dy.get_conn() as conn:
            self.assertTrue(dy._has_recent_regular_duty(conn.cursor(), self.first, '2026-03-29'))

    def test_current_manual_penalty_is_allowed_after_recent_regular_duty(self) -> None:
        self.history('2026-03-01')
        dy.insert_manual_override('2026-03-07', dy.GROUP_SEN, '00000001', '벌당직')
        with dy.get_conn() as conn:
            plan = dy.assign_day(conn.cursor(), datetime.date(2026, 3, 7))
            self.assertEqual(plan[0][1], '00000001')
            self.assertIsNone(conn.execute('SELECT emp_id FROM system_state WHERE pointer_key=?',
                                           (dy.PTR_SEN,)).fetchone())

    def test_return_queue_respects_regular_gap_and_penalty_exception(self) -> None:
        self.history('2026-03-01')
        with dy.get_conn() as conn:
            conn.execute('INSERT INTO skip_queue(emp_id,skipped_date,return_year_month) VALUES (?,?,?)',
                         (self.first, '2026-01-01', '2026-03'))
        self.assertEqual(self.choose('2026-03-20'), self.second)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT is_returned FROM skip_queue WHERE skipped_date=?',
                                           ('2026-01-01',)).fetchone()[0], 0)
        dy.insert_manual_override('2026-03-01', dy.GROUP_SEN, '00000001', '대근')
        self.assertEqual(self.choose('2026-03-21'), self.first)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT is_returned FROM skip_queue WHERE skipped_date=?',
                                           ('2026-01-01',)).fetchone()[0], 1)


if __name__ == '__main__':
    unittest.main()
