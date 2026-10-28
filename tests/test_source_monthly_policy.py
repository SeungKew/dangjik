"""원본 월별 정책의 경계값·복귀·원자성 및 7~10월 연속 회귀."""
from collections import Counter
import contextlib
import io
from pathlib import Path
import tempfile
import unittest

import dangjik_basic1 as dy
from app import app
from scripts.audit_monthly_policy import run_audit

ROOT = Path(__file__).resolve().parents[1]


class SourceMonthlyPolicyTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.original_db = dy.DB_NAME
        dy.DB_NAME = str(Path(self.temp.name) / 'test.db')
        dy.init_db()
        self.one = dy.register_employee('00000001', '첫직원', '부서', dy.GROUP_SUK)
        self.two = dy.register_employee('00000002', '다음직원', '부서', dy.GROUP_SUK)
        self.settings = {'mode':'source_monthly', 'prior_gap_days':30, 'current_month_only':True}

    def tearDown(self) -> None:
        dy.DB_NAME = self.original_db
        self.temp.cleanup()

    def duty(self, date: str, emp: int) -> None:
        with dy.get_conn() as conn:
            conn.execute('INSERT INTO duty_history(duty_date,duty_type,emp_id) VALUES (?,\'숙직\',?)', (date,emp))

    def test_previous_month_30_31_day_boundary_and_configurable_28(self) -> None:
        self.duty('2026-07-02', self.one)
        self.duty('2026-07-01', self.two)
        with dy.get_conn() as conn:
            self.assertEqual(dy._source_conflict(conn.cursor(),self.one,'2026-08-01',dy.PTR_REV,self.settings),('recent_gap','2026-07-02'))
            self.assertIsNone(dy._source_conflict(conn.cursor(),self.two,'2026-08-01',dy.PTR_REV,self.settings))
            self.assertIsNone(dy._source_conflict(conn.cursor(),self.one,'2026-08-01',dy.PTR_REV,{**self.settings,'prior_gap_days':28}))
            self.assertIsNone(dy._source_conflict(conn.cursor(),self.one,'2026-08-01',dy.PTR_JUN,self.settings))

    def test_future_weekday_reservation_wins_same_month_conflict(self) -> None:
        self.duty('2026-08-31',self.two)
        with dy.get_conn() as conn:
            self.assertEqual(dy._source_conflict(conn.cursor(),self.two,'2026-08-01',dy.PTR_REV,self.settings),('same_month','2026-08-31'))
            self.assertIsNone(dy._source_conflict(conn.cursor(),self.two,'2026-08-01',dy.PTR_REV,{**self.settings,'current_month_only':False}))
            chosen,_=dy._choose_source_monthly(conn.cursor(),'2026-08-01',dy.GROUP_SUK,dy.PTR_REV,self.settings)
            self.assertEqual(chosen,self.one)
            row=conn.execute('SELECT return_year_month,pointer_key,skip_reason,source_duty_date FROM skip_queue').fetchone()
            self.assertEqual(row,('2026-10',dy.PTR_REV,'same_month','2026-08-31'))

    def test_prior_month_skip_returns_from_original_duty_month(self) -> None:
        self.duty('2026-07-20',self.one)
        with dy.get_conn() as conn:
            chosen,_=dy._choose_source_monthly(conn.cursor(),'2026-08-03',dy.GROUP_SUK,dy.PTR_FWD,self.settings)
            self.assertEqual(chosen,self.two)
            self.assertEqual(conn.execute('SELECT return_year_month,skip_reason,source_duty_date FROM skip_queue').fetchone(),('2026-09','recent_gap','2026-07-20'))

    def test_current_return_batch_precedes_overdue_and_regular_pointer_stays(self) -> None:
        with dy.get_conn() as conn:
            for emp,month in ((self.one,'2026-09'),(self.two,'2026-10')):
                conn.execute('INSERT INTO skip_queue(emp_id,skipped_date,return_year_month,pointer_key) VALUES (?,\'2026-08-01\',?,?)',(emp,month,dy.PTR_FWD))
            chosen,_=dy._choose_source_monthly(conn.cursor(),'2026-10-01',dy.GROUP_SUK,dy.PTR_FWD,self.settings)
            self.assertEqual(chosen,self.two)
            self.assertIsNone(conn.execute('SELECT * FROM system_state WHERE pointer_key=?',(dy.PTR_FWD,)).fetchone())
            self.assertEqual(conn.execute('SELECT returned_date FROM skip_queue WHERE emp_id=?',(self.two,)).fetchone()[0],'2026-10-01')
            chosen,_=dy._choose_source_monthly(conn.cursor(),'2026-10-02',dy.GROUP_SUK,dy.PTR_FWD,self.settings)
            self.assertEqual(chosen,self.one)
            self.assertEqual(conn.execute('SELECT returned_date FROM skip_queue WHERE emp_id=?',(self.one,)).fetchone()[0],'2026-10-02')
        with contextlib.redirect_stdout(io.StringIO()):
            result=dy.delete_month('2026-10')
        self.assertEqual(result['returns_restored'],2)
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT sum(is_returned) FROM skip_queue').fetchone()[0],0)
            self.assertEqual(conn.execute('SELECT count(*) FROM skip_queue WHERE returned_date IS NOT NULL').fetchone()[0],0)

    def test_penalty_exception_and_manual_does_not_move_regular_pointer(self) -> None:
        dy.insert_manual_override('2026-07-31',dy.GROUP_SUK,'00000001','대근')
        self.duty('2026-07-31',self.one)
        with dy.get_conn() as conn:
            self.assertIsNone(dy._source_conflict(conn.cursor(),self.one,'2026-08-01',dy.PTR_REV,self.settings))
        dy.insert_manual_override('2026-08-03',dy.GROUP_SUK,'00000002','벌당직')
        with dy.get_conn() as conn:
            chosen,_=dy._choose_source_monthly(conn.cursor(),'2026-08-03',dy.GROUP_SUK,dy.PTR_FWD,self.settings)
            self.assertEqual(chosen,self.two)
            self.assertEqual(conn.execute('SELECT count(*) FROM system_state').fetchone()[0],0)

    def test_unsatisfiable_month_rolls_back_entire_generation(self) -> None:
        dy.set_generation_settings('source_monthly',30,True)
        with dy.get_conn() as conn:
            before=list(conn.iterdump())
        with contextlib.redirect_stdout(io.StringIO()),self.assertRaisesRegex(ValueError,'배정 대상자'):
            dy.regenerate_month('2026-08-01')
        with dy.get_conn() as conn:
            self.assertEqual(list(conn.iterdump()),before)

    def test_partial_or_past_month_request_leaves_existing_data_unchanged(self) -> None:
        dy.set_generation_settings('source_monthly',30,True)
        self.duty('2026-09-01',self.one)
        with dy.get_conn() as conn:
            before=list(conn.iterdump())
        for first,last in [('2026-08-02','2026-08-31'),('2026-08-01','2026-08-31')]:
            with self.assertRaises(ValueError):
                dy.generate_duty_plan(first,last)
            with dy.get_conn() as conn:
                self.assertEqual(list(conn.iterdump()),before)

    @unittest.skipUnless((ROOT/'docs/2026년_8월_본사_당직명령.xlsx').exists(),'로컬 원본 필요')
    def test_web_roster_only_updates_august_staff_without_loading_answers(self) -> None:
        from unittest.mock import patch
        import app as web
        dy.load_original_order(str(ROOT/'docs/2026년_7월_본사_당직명령.xlsx'))
        with dy.get_conn() as conn:
            before={t:conn.execute(f'SELECT * FROM {t} ORDER BY rowid').fetchall() for t in ('duty_history','skip_queue','system_state','system_state_snapshot','manual_overrides','system_settings')}
        with patch.object(web,'UPLOAD_DIR',self.temp.name):
            response=app.test_client().post('/upload',data={'import_scope':'roster','file':(io.BytesIO((ROOT/'docs/2026년_8월_본사_당직명령.xlsx').read_bytes()),'8월.xlsx')},content_type='multipart/form-data',follow_redirects=True)
        self.assertIn('직원 명부만 갱신했습니다',response.get_data(as_text=True))
        with dy.get_conn() as conn:
            for table,rows in before.items():
                self.assertEqual(conn.execute(f'SELECT * FROM {table} ORDER BY rowid').fetchall(),rows)
            self.assertEqual(conn.execute("SELECT is_active FROM employees WHERE emp_no='17161326'").fetchone()[0],1)
            self.assertEqual(conn.execute("SELECT is_active FROM employees WHERE emp_no='12161921'").fetchone()[0],0)
            self.assertEqual(conn.execute("SELECT count(*) FROM duty_history WHERE duty_date LIKE '2026-08%'").fetchone()[0],0)

    @unittest.skipUnless((ROOT/'docs/2026년_10월_본사_당직명령.xlsx').exists(),'로컬 원본 필요')
    def test_original_without_skip_table_cannot_erase_pending_returns(self) -> None:
        with dy.get_conn() as conn:
            before=list(conn.iterdump())
        path=str(ROOT/'docs/2026년_10월_본사_당직명령.xlsx')
        with self.assertRaisesRegex(ValueError,'복귀 대기'):
            dy.load_original_order(path)
        with dy.get_conn() as conn:
            self.assertEqual(list(conn.iterdump()),before)
        result=dy.load_original_order(path,roster_only=True)
        self.assertEqual(result['month'],'2026-10')
        self.assertTrue(any('스킵 내역 표' in w for w in result['warnings']))

    def test_web_generation_failure_preserves_history_and_reports_reason(self) -> None:
        dy.set_generation_settings('source_monthly',30,True)
        self.duty('2026-09-01',self.one)
        with dy.get_conn() as conn:
            before=list(conn.iterdump())
        response=app.test_client().post('/generate',data={'month':'2026-08'},follow_redirects=True)
        self.assertEqual(response.status_code,200)
        self.assertIn('이후 월을 먼저 삭제',response.get_data(as_text=True))
        with dy.get_conn() as conn:
            self.assertEqual(list(conn.iterdump()),before)
        response=app.test_client().post('/manual',data={'target_date':'2026-08-01','emp_no':'00000001','reason':'대근','regen':'on'},follow_redirects=True)
        self.assertEqual(response.status_code,200)
        self.assertIn('수동 예외는 저장됐지만',response.get_data(as_text=True))
        with dy.get_conn() as conn:
            self.assertEqual(conn.execute('SELECT count(*) FROM manual_overrides').fetchone()[0],1)
            self.assertEqual(conn.execute('SELECT count(*) FROM duty_history').fetchone()[0],1)

    def test_admin_settings_and_invalid_input_are_atomic(self) -> None:
        client=app.test_client()
        result=client.post('/admin/skip-policy',data={'generation_mode':'source_monthly','prior_gap_days':'30','skip_current_month_only':'on'},follow_redirects=True)
        self.assertEqual(result.status_code,200)
        self.assertEqual(dy.get_generation_settings(),self.settings)
        self.assertIn('30일',result.get_data(as_text=True))
        client.post('/admin/skip-policy',data={'generation_mode':'legacy','prior_gap_days':'-1'})
        self.assertEqual(dy.get_generation_settings(),self.settings)
        dy.set_generation_settings('legacy',28,False)
        self.assertEqual(dy.get_generation_settings()['mode'],'legacy')

    @unittest.skipUnless((ROOT/'docs/2026년_10월_본사_당직명령.xlsx').exists(),'로컬 원본 필요')
    def test_sequential_july_to_october_has_no_answer_history_or_queue_import(self) -> None:
        report=run_audit(30)
        self.assertEqual([m['matched'] for m in report['months']],[49,53,50,53])
        self.assertEqual(report['total_matched'],205)
        self.assertTrue(report['protected_files_unchanged'])
        august=report['months'][1]
        reasons=Counter((x[4],x[3]) for x in august['generated_skips'])
        self.assertEqual(reasons,{('same_month','2026-10'):11,('recent_gap','2026-09'):23})
        self.assertEqual(len(report['months'][2]['pending']),22)
        self.assertEqual(len(report['months'][3]['pending']),11)

    @unittest.skipUnless((ROOT/'docs/2026년_8월_본사_당직명령.xlsx').exists(),'로컬 원본 필요')
    def test_regeneration_and_roster_only_import_preserve_results_and_queue(self) -> None:
        dy.load_original_order(str(ROOT/'docs/2026년_7월_본사_당직명령.xlsx'))
        dy.load_original_order(str(ROOT/'docs/2026년_8월_본사_당직명령.xlsx'),roster_only=True)
        def state() -> tuple:
            with dy.get_conn() as conn:
                duties=conn.execute('SELECT duty_date,duty_type,emp_id FROM duty_history ORDER BY duty_date,duty_type,emp_id').fetchall()
                queue=conn.execute('SELECT emp_id,skipped_date,return_year_month,pointer_key,is_returned,returned_date FROM skip_queue ORDER BY skipped_date,pointer_key,emp_id').fetchall()
                return duties,queue
        with contextlib.redirect_stdout(io.StringIO()):
            dy.regenerate_month('2026-08-01')
            first=state()
            dy.regenerate_month('2026-08-01')
            self.assertEqual(state(),first)
            dy.delete_month('2026-08')
            dy.regenerate_month('2026-08-01')
            self.assertEqual(state(),first)
        # 9월 원본에서 남아 있는 9월 복귀 11명은 아직 미소진이다.
        source=dy.read_original_order(str(ROOT/'docs/2026년_9월_본사_당직명령.xlsx'))
        current=[s for s in source['skips'] if s['return_year_month']=='2026-09']
        self.assertEqual(len(current),11)
        self.assertTrue(all(not x['is_returned'] for x in current))


if __name__=='__main__':
    unittest.main()
