"""7~10월 원본 기준 월별 배정 회귀. 임시 DB에서만 실행한다.

기준 이력은 6월 원본, 이후 월별로 명부만 갱신해 다음 달 정답 이력/큐를
사용하지 않고 연속 생성한다. 원본 7/22 동명이인 사번 보완을 별도 기록한다.
"""
from collections import Counter
import contextlib
import hashlib
import io
import json
from pathlib import Path
import sqlite3
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import dangjik_basic1 as dy
from scripts.compare_april import read_order


def run_audit(gap_days: int = 30) -> dict:
    originals = {m: ROOT / f'docs/2026년_{m}월_본사_당직명령.xlsx' for m in range(6, 11)}
    protected = [*originals.values(), Path(dy.DB_NAME), Path(dy.V2_DB_NAME)]
    hashes = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in protected if p.exists()}
    previous_db = dy.DB_NAME
    results: list[dict] = []
    try:
        with tempfile.TemporaryDirectory() as temp:
            dy.DB_NAME = str(Path(temp) / 'audit.db')
            dy.load_original_order(str(originals[6]))
            dy.set_generation_settings('source_monthly', gap_days, True)
            for month in range(7, 11):
                source = dy.read_original_order(str(originals[month]))
                with dy.get_conn() as conn:
                    history_before = conn.execute('SELECT * FROM duty_history ORDER BY id').fetchall()
                    queue_before = conn.execute('SELECT * FROM skip_queue ORDER BY id').fetchall()
                    state_before = conn.execute('SELECT * FROM system_state ORDER BY pointer_key').fetchall()
                # DB에 목표 월 정답 이력과 스킵 큐를 쓰지 않는다.
                dy.load_original_order(str(originals[month]), roster_only=True)
                with dy.get_conn() as conn:
                    assert conn.execute('SELECT * FROM duty_history ORDER BY id').fetchall() == history_before
                    assert conn.execute('SELECT * FROM skip_queue ORDER BY id').fetchall() == queue_before
                    assert conn.execute('SELECT * FROM system_state ORDER BY pointer_key').fetchall() == state_before
                    assert conn.execute('SELECT count(*) FROM duty_history WHERE duty_date LIKE ?', (source['month']+'%',)).fetchone()[0] == 0
                with contextlib.redirect_stdout(io.StringIO()):
                    dy.regenerate_month(source['month'] + '-01')
                with dy.get_conn() as conn:
                    got = Counter(conn.execute('''SELECT h.duty_date,h.duty_type,e.emp_no FROM duty_history h
                        JOIN employees e ON e.id=h.emp_id WHERE h.duty_date LIKE ?''', (source['month']+'%',)).fetchall())
                    skips = conn.execute('''SELECT e.emp_no,e.name,q.pointer_key,q.return_year_month,q.skip_reason,q.source_duty_date
                        FROM skip_queue q JOIN employees e ON e.id=q.emp_id
                        WHERE q.skipped_date LIKE ? ORDER BY q.id''', (source['month']+'%',)).fetchall()
                    pending = conn.execute('''SELECT e.emp_no,e.name,q.pointer_key,q.return_year_month
                        FROM skip_queue q JOIN employees e ON e.id=q.emp_id WHERE q.is_returned=0
                        ORDER BY q.return_year_month DESC,q.skipped_date,q.id''').fetchall()
                expected = Counter((d['date'],d['kind'],d['emp_no']) for d in source['duties'])
                raw_expected = Counter((d['date'],d['kind'],d['emp_no']) for d in read_order(originals[month],month))
                source_pending = Counter((s['emp_no'],s['pointer_key'],s['return_year_month']) for s in source['skips'] if not s['is_returned'])
                actual_pending = Counter((s[0],s[2],s[3]) for s in pending)
                result = dict(month=source['month'], raw_matched=sum((got & raw_expected).values()),
                              pending_matches_source_table=actual_pending == source_pending if month != 10 else None,
                              expected=sum(expected.values()), generated=sum(got.values()),
                              matched=sum((got & expected).values()), missing=list((expected-got).elements()),
                              extra=list((got-expected).elements()), generated_skips=skips, pending=pending,
                              source_warnings=source['warnings'], roster_import_preserved_history_and_queue=True)
                results.append(result)
    finally:
        dy.DB_NAME = previous_db
    return dict(gap_days=gap_days, months=results, total_matched=sum(x['matched'] for x in results),
                total_expected=sum(x['expected'] for x in results),
                protected_files_unchanged=all(hashlib.sha256(p.read_bytes()).hexdigest()==v for p,v in hashes.items()))


def main() -> None:
    out = ROOT / 'audit_results/policy_2026_07_10'
    out.mkdir(parents=True, exist_ok=True)
    report = run_audit()
    (out / 'generation_verification.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n')
    print(json.dumps({**{k:v for k,v in report.items() if k!='months'},
                      'months':[{k:v for k,v in m.items() if k not in ('generated_skips','pending')} for m in report['months']]}, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
