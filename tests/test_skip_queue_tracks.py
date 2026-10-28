"""평일/휴일 숙직의 스킵·복귀 트랙 보존 및 구 큐 이관 검증."""

from pathlib import Path
import tempfile
import unittest

import dangjik_basic1 as dy


class SkipQueueTracksTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.original_db = dy.DB_NAME
        dy.DB_NAME = str(Path(self.tmp.name) / 'queue.db')
        dy.init_db()
        self.first = dy.register_employee('00000001', '첫직원', '부서', dy.GROUP_SUK)
        self.second = dy.register_employee('00000002', '다음직원', '부서', dy.GROUP_SUK)

    def tearDown(self) -> None:
        dy.DB_NAME = self.original_db
        self.tmp.cleanup()

    def queue(self, emp_id: int, date: str, track: str | None) -> None:
        with dy.get_conn() as conn:
            conn.execute('INSERT INTO skip_queue(emp_id,skipped_date,return_year_month,pointer_key) '
                         'VALUES (?,?,?,?)', (emp_id, date, '2026-06', track))

    def choose(self, date: str, reverse: bool) -> int | None:
        with dy.get_conn() as conn:
            emp, _ = dy.get_next_candidate(conn.cursor(), date, dy.GROUP_SUK,
                                           dy.PTR_REV if reverse else dy.PTR_FWD, reverse)
            return emp

    def test_weekday_queue_is_not_consumed_by_holiday(self) -> None:
        self.queue(self.first, '2026-04-28', dy.PTR_FWD)
        self.assertEqual(self.choose('2026-06-03', True), self.second)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT is_returned FROM skip_queue').fetchone()[0], 0)
        self.assertEqual(self.choose('2026-06-04', False), self.first)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT is_returned FROM skip_queue').fetchone()[0], 1)
            self.assertIsNone(conn.execute('SELECT emp_id FROM system_state WHERE pointer_key=?',
                                           (dy.PTR_FWD,)).fetchone())

    def test_holiday_queue_is_not_consumed_by_weekday(self) -> None:
        self.queue(self.second, '2026-04-25', dy.PTR_REV)
        self.assertEqual(self.choose('2026-06-04', False), self.first)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT is_returned FROM skip_queue').fetchone()[0], 0)
        self.assertEqual(self.choose('2026-06-06', True), self.second)

    def test_legacy_schema_preserves_rows_and_backfills_idempotently(self) -> None:
        with dy.get_conn() as conn:
            conn.execute('DROP TABLE skip_queue')
            conn.execute('''CREATE TABLE skip_queue (
                id INTEGER PRIMARY KEY AUTOINCREMENT, emp_id INTEGER, skipped_date DATE,
                return_year_month TEXT, is_returned INTEGER DEFAULT 0)''')
            conn.execute('INSERT INTO skip_queue(emp_id,skipped_date,return_year_month,is_returned) VALUES (?,?,?,?)',
                         (self.first, '2026-04-28', '2026-06', 1))
            conn.execute('INSERT INTO skip_queue(emp_id,skipped_date,return_year_month,is_returned) VALUES (?,?,?,?)',
                         (self.second, '2026-06-03', '2026-08', 0))
            before = conn.execute('SELECT * FROM skip_queue ORDER BY id').fetchall()
        dy.init_db()
        dy.init_db()
        with dy.get_conn() as conn:
            after = conn.execute('SELECT * FROM skip_queue ORDER BY id').fetchall()
        self.assertEqual([r[:5] for r in after], before)
        self.assertEqual([r[5] for r in after], [dy.PTR_FWD, dy.PTR_REV])

    def test_missing_track_after_initialization_uses_legacy_fallback(self) -> None:
        self.queue(self.first, '2026-04-28', None)
        self.assertEqual(self.choose('2026-06-03', True), self.second)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT pointer_key,is_returned FROM skip_queue').fetchone(),
                             (dy.PTR_FWD, 0))

    def test_new_skips_store_each_original_pointer(self) -> None:
        with dy.get_conn() as conn:
            conn.execute('INSERT INTO duty_history(duty_date,duty_type,emp_id) VALUES (?,?,?)',
                         ('2026-05-01', '숙직', self.first))
        self.assertEqual(self.choose('2026-05-20', False), self.second)
        with dy.get_conn() as conn:
            conn.execute('REPLACE INTO system_state VALUES (?,?)', (dy.PTR_REV, self.second))
        self.assertEqual(self.choose('2026-05-23', True), self.second)
        with dy.get_conn() as conn:
            rows = conn.execute('SELECT pointer_key,return_year_month FROM skip_queue ORDER BY id').fetchall()
        self.assertEqual(rows, [(dy.PTR_FWD, '2026-07'), (dy.PTR_REV, '2026-07')])

    def test_explicit_track_survives_later_master_change(self) -> None:
        self.queue(self.first, '2026-04-28', dy.PTR_FWD)
        with dy.get_conn() as conn:
            conn.execute('UPDATE employees SET duty_group=? WHERE id=?', (dy.GROUP_JUN, self.first))
        dy.init_db()
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT pointer_key FROM skip_queue').fetchone()[0], dy.PTR_FWD)


if __name__ == '__main__':
    unittest.main()
