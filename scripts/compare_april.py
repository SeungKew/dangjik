"""Compare the generated April assignments in the production DB against the
original April workbook in docs/, without changing production data.

Read-only comparisons:
  .venv/bin/python scripts/compare_april.py
Results are written under audit_results/2026-04.
"""

import csv
import hashlib
import json
from pathlib import Path
import sqlite3

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / 'data/duty_system_v3.db'
SOURCE = ROOT / 'docs/2026년_4월_본사_당직명령.xlsx'
OUT = ROOT / 'audit_results/2026-04'
MONTH = 4


def fingerprint(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_order(path: Path, month: int) -> list[dict]:
    """Read the 당직명령 sheet (columns B:G only). Same parser as audit_march.py."""
    wb = openpyxl.load_workbook(path, data_only=True)
    rows: list[dict] = []
    day, kind, slot = 0, '', 0
    for row in wb['당직명령'].iter_rows(min_row=4):
        if isinstance(row[1].value, (int, float)):
            day, slot = int(row[1].value), 0
        if row[3].value in ('일직', '숙직'):
            kind = row[3].value
        if row[5].value is None or row[6].value is None:
            continue
        emp_no = str(int(row[6].value)) if isinstance(row[6].value, float) else str(row[6].value)
        rows.append(dict(date=f'2026-{month:02d}-{day:02d}', kind=kind,
                         name=str(row[5].value).strip(), emp_no=emp_no.zfill(8),
                         penalty=bool(row[5].font.bold)))
        slot += 1
    wb.close()
    return rows


def read_db(db_path: Path) -> tuple[list[dict], dict]:
    conn = sqlite3.connect(f'file:{db_path}?mode=ro', uri=True)
    try:
        rows = conn.execute(
            "SELECT h.duty_date,h.duty_type,e.emp_no,e.name FROM duty_history h "
            "JOIN employees e ON h.emp_id=e.id "
            f"WHERE h.duty_date LIKE '2026-{MONTH:02d}%' ORDER BY h.duty_date,h.id"
        ).fetchall()
        master = {str(n).zfill(8): (name, group, active) for n, name, group, active in conn.execute(
            'SELECT emp_no,name,duty_group,is_active FROM employees')}
    finally:
        conn.close()
    return [dict(date=d, kind=k, emp_no=str(n).zfill(8), name=name) for d, k, n, name in rows], master


def main() -> None:
    protected = [SOURCE, DB, ROOT / 'data/duty_system_v2.db']
    before = {p.name: fingerprint(p) for p in protected}

    actual = read_order(SOURCE, MONTH)
    generated, master = read_db(DB)

    actual_keys = {(r['date'], r['kind'], r['emp_no']) for r in actual}
    generated_keys = {(r['date'], r['kind'], r['emp_no']) for r in generated}
    generated_by_key = {(r['date'], r['kind'], r['emp_no']): r for r in generated}

    matched = [r for r in actual if (r['date'], r['kind'], r['emp_no']) in generated_keys]
    missing = [r for r in actual if (r['date'], r['kind'], r['emp_no']) not in generated_keys]
    extra = [r for r in generated if (r['date'], r['kind'], r['emp_no']) not in actual_keys]

    def master_state(emp_no: str) -> str:
        state = master.get(emp_no)
        if state is None:
            return '명부없음'
        return '활성' if state[2] else '제외'

    summary = {
        'source_file': SOURCE.name,
        'actual_count': len(actual),
        'generated_count': len(generated),
        'matched_count': len(matched),
        'exact_name_match_count': sum(
            generated_by_key[(r['date'], r['kind'], r['emp_no'])]['name'] == r['name'] for r in matched),
        'mismatched_count': len(actual) - len(matched),
        'match_rate': f"{len(matched)}/{len(actual)}",
        'penalty_count': sum(r['penalty'] for r in actual),
        'missing_from_generated': [
            {**r, '마스터상태': master_state(r['emp_no'])} for r in missing],
        'extra_in_generated': extra,
        'name_conflicts': [
            {'date': a['date'], 'kind': a['kind'], 'emp_no': a['emp_no'],
             '원본이름': a['name'],
             'DB이름': generated_by_key[(a['date'], a['kind'], a['emp_no'])]['name']}
            for a in matched
            if a['name'] != generated_by_key[(a['date'], a['kind'], a['emp_no'])]['name']
        ][:50],
        'production_hashes_unchanged': all(
            fingerprint(p) == before[p.name] for p in protected),
    }

    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / 'latest_comparison.json').write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + '\n')

    with (OUT / 'latest_comparison.csv').open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['날짜', '분류', '원본_이름', '원본_사번', '벌당직', 'DB_이름', 'DB_사번', '결과', '마스터상태'])
        gen_by_slot = {}
        for r in generated:
            gen_by_slot.setdefault((r['date'], r['kind']), []).append(r)
        actual_by_slot = {}
        for r in actual:
            actual_by_slot.setdefault((r['date'], r['kind']), []).append(r)
        for r in actual:
            slot_rows = gen_by_slot.get((r['date'], r['kind']), [])
            g = next((x for x in slot_rows if x['emp_no'] == r['emp_no']), None)
            if g is not None:
                result = '일치' if g['name'] == r['name'] else '사번일치·이름상이'
                g_name, g_no = g['name'], g['emp_no']
            else:
                result, g_name, g_no = '불일치', '', ''
            writer.writerow([r['date'], r['kind'], r['name'], r['emp_no'], r['penalty'],
                             g_name, g_no, result, master_state(r['emp_no'])])
        for r in extra:
            writer.writerow([r['date'], r['kind'], '', '', False, r['name'], r['emp_no'],
                             'DB에만있음', master_state(r['emp_no'])])

    print(json.dumps({k: v for k, v in summary.items() if k != 'extra_in_generated'},
                     ensure_ascii=False, indent=2))
    if extra:
        print('extra_in_generated:', json.dumps(extra, ensure_ascii=False))
    print(f'output: {OUT}')


if __name__ == '__main__':
    main()
