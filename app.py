"""본사 당직명령 배정 시스템 — Flask 웹 앱.

당직 배정 로직은 dangjik_basic1.py를 그대로 재사용한다(검증 완료 재현 규칙 기반).
기능:
  1. 입력양식 엑셀(xlsx) 업로드 → 직원 마스터 동기화 + 초기당직이력 적재
  2. 배정 생성 (기간 지정)
  3. 수동 예외 배정(벌당직/대근) 등록·취소 후 해당 월 재배정 반영
  4. 직원 변동 처리 (퇴사=제외, 승진=그룹/순번 변경)
"""

import datetime
import os
import tempfile
from io import BytesIO

from flask import Flask, Response, flash, redirect, render_template, request, send_file, url_for

import dangjik_basic1 as dy

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')

app = Flask(__name__)
app.secret_key = os.environ.get('DANGJIK_SECRET_KEY', 'dangjik-local-secret')


@app.context_processor
def inject_nav() -> dict:
    nav = [
        ('dashboard', '배정 현황'),
        ('upload', '입력양식 등록'),
        ('generate', '배정 생성'),
        ('manual', '수동 예외 배정'),
        ('employees', '직원 관리'),
        ('employee_history', '직원 변동 이력'),
        ('admin', '시스템 관리'),
    ]
    return {'nav_items': nav}


@app.route('/')
def dashboard() -> str:
    year_month = request.args.get('month') or datetime.date.today().strftime('%Y-%m')
    if not _is_valid_ym(year_month):
        flash('올바르지 않은 연월입니다 (YYYY-MM).', 'error')
        year_month = datetime.date.today().strftime('%Y-%m')
    rows = [decorate_assignment(r) for r in dy.get_assignments(year_month)]
    counts = {
        'employees': _count('employees'),
        'history': _count('duty_history'),
        'skipped': _count('skip_queue', 'is_returned = 0'),
        'overrides': _count('manual_overrides'),
    }
    return render_template('dashboard.html', rows=rows, year_month=year_month, counts=counts)


@app.route('/export/excel')
def download_assignments_excel() -> Response:
    year_month = request.args.get('month', '')
    if not _is_valid_ym(year_month):
        flash('올바르지 않은 연월입니다 (YYYY-MM).', 'error')
        return redirect(url_for('dashboard'))
    try:
        content = dy.export_assignments_excel(year_month)
    except ValueError as exc:
        flash(str(exc), 'error')
        return redirect(url_for('dashboard', month=year_month))
    return send_file(
        BytesIO(content), as_attachment=True,
        download_name=f'{year_month}_본사_당직명령.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


_HOL_NAME_KR = {
    "New Year's Day": '신정',
    "The day preceding Korean New Year": '설날 연휴',
    "Korean New Year": '설날',
    "The second day of Korean New Year": '설날 연휴',
    "Independence Movement Day": '삼일절',
    "Children's Day": '어린이날',
    "Buddha's Birthday": '부처님 오신 날',
    "Memorial Day": '현충일',
    "Liberation Day": '광복절',
    "National Foundation Day": '개천절',
    "Hangul Day": '한글날',
    "Christmas Day": '성탄절',
}

_hol_cache: dict[int, dict] = {}


def _kr_calendar(year: int) -> dict:
    """연도별 한국 공휴일 달력 캐시 (행마다 재생성 방지)."""
    if year not in _hol_cache:
        _hol_cache[year] = dy.holidays.KR(years=[year])
    return _hol_cache[year]


def decorate_assignment(row: tuple) -> dict:
    """(날짜, 분류, 이름, 사번, 그룹, 종류) → 화면 표시용 dict."""
    duty_date, duty_type, name, emp_no, group, kind = row
    d = datetime.datetime.strptime(duty_date[:10], '%Y-%m-%d').date()
    hol_name = _kr_calendar(d.year).get(d)  # 공휴일 아니면 None
    return {
        'date': duty_date[:10],
        'dow': '월화수목금토일'[d.weekday()],
        'holiday_name': _HOL_NAME_KR.get(hol_name, '공휴일') if hol_name else '',
        'day_class': ('weekend' if d.weekday() >= 5 else 'holiday') if hol_name or d.weekday() >= 5 else '',
        'duty_type': duty_type,
        'name': name,
        'emp_no': emp_no,
        'group': group,
        'kind': kind if kind != '배정' else '',
    }


def _is_valid_ym(year_month: str) -> bool:
    try:
        datetime.datetime.strptime(year_month, '%Y-%m')
        return True
    except ValueError:
        return False


def _count(table: str, extra_where: str = '') -> int:
    conn = dy.get_conn()
    c = conn.cursor()
    c.execute(f"SELECT COUNT(*) FROM {table}{' WHERE ' + extra_where if extra_where else ''}")
    n = int(c.fetchone()[0])
    conn.close()
    return n


@app.route('/upload', methods=['GET', 'POST'])
def upload() -> str:
    if request.method == 'POST':
        file = request.files.get('file')
        if file is None or not file.filename:
            flash('업로드할 파일을 선택하세요.', 'error')
            return redirect(url_for('upload'))
        if not file.filename.lower().endswith('.xlsx'):
            flash('원본 당직명령 또는 기존 입력양식의 .xlsx 파일을 선택하세요. .xls는 .xlsx로 변환해 주세요.', 'error')
            return redirect(url_for('upload'))
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=UPLOAD_DIR, suffix='.xlsx', delete=False) as staged:
            file_path = staged.name
        try:
            file.save(file_path)
            result = dy.import_uploaded_workbook(file_path, roster_only=request.form.get('import_scope') == 'roster')
        except Exception as exc:  # noqa: BLE001 — 업로드 파일 형식 오류를 사용자에게 안내
            flash(f'업로드 처리 실패: {exc}', 'error')
            return redirect(url_for('upload'))
        finally:
            os.remove(file_path)
        if result['format'] == 'original_roster':
            flash(f"원본 {result['month']} 직원 명부만 갱신했습니다: {result['employees']}명. "
                  f"기존 당직 이력·스킵·시작 순번은 보존했습니다. 백업: {result['backup']}", 'success')
        elif result['format'] == 'original':
            flash(f"원본 {result['month']} 기준 등록 완료: 직원 {result['employees']}명, "
                  f"당직 {result['duties']}건, 스킵 {result['skips']}건. 백업: {result['backup']}", 'success')
        else:
            flash(f'입력양식 등록 완료: {file.filename}', 'success')
        for warning in result['warnings']:
            flash(warning, 'info')
        return redirect(url_for('upload'))
    return render_template('upload.html')


@app.route('/generate', methods=['GET', 'POST'])
def generate() -> str:
    """월 단위 배정 생성. 배정은 실제로 월 단위로 생성되므로 연월(YYYY-MM)만 지정한다."""
    if request.method == 'POST':
        year_month = request.form.get('month', '')
        if not _is_valid_ym(year_month):
            flash('연월 형식은 YYYY-MM 입니다.', 'error')
            return redirect(url_for('generate'))
        try:
            dy.regenerate_month(year_month + '-01')
        except ValueError as exc:
            flash(f'배정 생성 실패: {exc}', 'error')
            return redirect(url_for('generate'))
        flash(f'{year_month} 배정 생성 완료.', 'success')
        return redirect(url_for('dashboard', month=year_month))
    year_month = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime('%Y-%m')
    return render_template('generate.html', year_month=year_month)


@app.route('/manual', methods=['GET', 'POST'])
def manual() -> str:
    if request.method == 'POST':
        target_date = request.form['target_date']
        emp_no = request.form['emp_no']
        reason = request.form['reason'].strip()
        try:
            datetime.datetime.strptime(target_date, '%Y-%m-%d')
        except ValueError:
            flash('날짜 형식은 YYYY-MM-DD 입니다.', 'error')
            return redirect(url_for('manual'))
        if not reason:
            flash('사유(벌당직, 대근 등)를 입력하세요.', 'error')
            return redirect(url_for('manual'))
        emp = _find_emp(emp_no)
        if emp is None:
            flash(f'해당 사번을 찾을 수 없습니다: {emp_no}', 'error')
            return redirect(url_for('manual'))
        ok = dy.insert_manual_override(target_date, emp['group'], emp['no'], reason)
        flash(f'수동 배정 등록 {"완료" if ok else "실패"}: {target_date} {emp["name"]}({emp["no"]}) {reason}',
              'success' if ok else 'error')
        if request.form.get('regen') == 'on':
            try:
                dy.regenerate_month(target_date)
            except ValueError as exc:
                flash(f'수동 예외는 저장됐지만 재배정하지 못했습니다: {exc}', 'error')
            else:
                flash(f'{target_date[:7]} 전체 재배정 완료 (수동 배정 반영)', 'success')
        return redirect(url_for('manual'))
    overrides = _list_overrides()
    employees = [ '|'.join((e[1], str(e[2]), e[3], str(e[4]))) for e in _list_active_for_select() ]
    employees.sort(key=lambda s: (s.split('|')[0], int(s.split('|')[1])))
    return render_template('manual.html', overrides=overrides, employee_options=employees)


def _find_emp(emp_no: str) -> dict | None:
    emp_no = str(emp_no).strip().zfill(8)
    conn = dy.get_conn()
    c = conn.cursor()
    c.execute("SELECT emp_no, name, duty_group, group_idx FROM employees WHERE emp_no = ?", (emp_no,))
    row = c.fetchone()
    conn.close()
    return {'no': row[0], 'name': row[1], 'group': row[2], 'idx': row[3]} if row else None


def _list_active_for_select() -> list[tuple]:
    conn = dy.get_conn()
    c = conn.cursor()
    c.execute('''SELECT id, duty_group, group_idx, name, emp_no FROM employees
                 WHERE is_active = 1 ORDER BY duty_group, group_idx''')
    rows = c.fetchall()
    conn.close()
    return rows


def _list_overrides() -> list[dict]:
    conn = dy.get_conn()
    c = conn.cursor()
    c.execute('''SELECT m.id, m.target_date, m.duty_group, m.emp_no, e.name, m.reason
                 FROM manual_overrides m LEFT JOIN employees e ON m.emp_no = e.emp_no
                 ORDER BY m.target_date DESC, m.id DESC''')
    rows = c.fetchall()
    conn.close()
    return [dict(id=r[0], date=r[1], group=r[2], emp_no=r[3],
                 name=(r[4] if r[4] else '(사번 미일치)'), reason=r[5]) for r in rows]


@app.route('/manual/<int:override_id>/delete', methods=['POST'])
def manual_delete(override_id: int) -> str:
    dy.delete_manual_override(override_id)
    flash('수동 배정을 취소했습니다.', 'success')
    return redirect(url_for('manual'))


@app.route('/regenerate', methods=['POST'])
def regenerate() -> str:
    year_month = request.form.get('month', '')
    if not _is_valid_ym(year_month):
        flash('올바르지 않은 연월입니다.', 'error')
        return redirect(url_for('dashboard'))
    dy.regenerate_month(year_month + '-01')
    flash(f'{year_month} 전체 재배정 완료.', 'success')
    return redirect(url_for('dashboard', month=year_month))


@app.route('/employees')
def employees() -> str:
    group = request.args.get('group', 'all')
    status = request.args.get('status', 'all')
    keyword = request.args.get('q', '').strip() or None
    rows = dy.list_employees(duty_group=group, status=status, keyword=keyword)
    groups = sorted(set(dy.get_duty_groups()) | set(dy.EMPLOYEE_GROUPS))
    return render_template('employees.html', rows=rows, groups=groups,
                           cur_group=group, cur_status=status, cur_q=keyword or '',
                           change_date=dy.employee_change_date())


@app.route('/employees/register', methods=['POST'])
def employee_register() -> str:
    try:
        dy.register_employee(request.form.get('emp_no', ''), request.form.get('name', ''),
                             request.form.get('dept', ''), request.form.get('duty_group', ''),
                             reason=request.form.get('reason', ''), effective_date=request.form.get('effective_date'))
    except ValueError as exc:
        flash(str(exc), 'error')
    else:
        flash('전입 직원을 등록했습니다. 해당 그룹의 마지막 순번부터 배정 대상에 포함됩니다.', 'success')
    return redirect(url_for('employees'))


@app.route('/employees/<int:emp_id>/transfer-out', methods=['POST'])
def employee_transfer_out(emp_id: int) -> str:
    try:
        ok = dy.transfer_out_employee(emp_id, request.form.get('reason', ''), request.form.get('effective_date'))
    except ValueError as exc:
        flash(str(exc), 'error')
        return redirect(url_for('employees'))
    if ok:
        flash('전출 처리했습니다. 이후 정규 배정에서 제외되며 기존 이력과 수동 배정은 보존됩니다.', 'success')
    else:
        flash('해당 직원이 존재하지 않습니다.', 'error')
    return redirect(url_for('employees'))


@app.route('/employees/<int:emp_id>/update', methods=['POST'])
def employee_update(emp_id: int) -> str:
    form = request.form
    try:
        group_idx = int(form['group_idx'])
        ok = dy.update_employee(emp_id,
                            duty_group=form['duty_group'],
                            group_idx=group_idx,
                            name=form['name'],
                            dept=form['dept'],
                            is_active=1 if form.get('is_active_sel', '0') == '1' else 0,
                            reason=form.get('reason', ''), effective_date=form.get('effective_date'))
    except KeyError:
        flash('직원 정보의 필수 항목을 입력하세요.', 'error')
        return redirect(url_for('employees'))
    except ValueError as exc:
        flash('순번은 숫자이며 변동일은 올바른 날짜여야 합니다. ' + str(exc), 'error')
        return redirect(url_for('employees'))
    msg = '직원 정보를 갱신했습니다.' if ok else f'해당 직원(id={emp_id})이 존재하지 않습니다.'
    flash(msg, 'success' if ok else 'error')
    if ok:
        flash('변동 내역 반영을 원하면 "배정 현황"에서 해당 월 전체 재배정을 실행하세요.', 'info')
    return redirect(url_for('employees', group=form.get('duty_group', 'all'),
                            status=form.get('status', 'all')))


@app.route('/employees/history')
def employee_history() -> str:
    emp_id = request.args.get('emp_id', type=int)
    keyword = request.args.get('q', '').strip()
    change_type = request.args.get('type', '')
    page = max(1, request.args.get('page', default=1, type=int))
    page_size = 50
    rows, total = dy.list_employee_changes(keyword, emp_id, change_type, page=page, page_size=page_size)
    pages = max(1, (total + page_size - 1) // page_size)
    page = min(page, pages)
    return render_template('employee_history.html', rows=rows, total=total, emp_id=emp_id,
                           keyword=keyword, change_type=change_type, page=page, pages=pages,
                           first=(page - 1) * page_size + 1 if total else 0,
                           last=min(page * page_size, total))


@app.template_filter('history_time')
def format_history_time(value: str) -> str:
    """저장된 이력 시각을 한국 시간의 날짜·시간으로 표시한다."""
    if not value:
        return '—'
    try:
        stamp = datetime.datetime.fromisoformat(value)
    except ValueError:
        return '—'
    kst = datetime.timezone(datetime.timedelta(hours=9))
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=kst)
    return stamp.astimezone(kst).strftime('%Y-%m-%d %H시')


@app.route('/admin')
def admin() -> str:
    counts = {
        'employees': _count('employees'),
        'active': _count('employees', 'is_active = 1'),
        'history': _count('duty_history'),
        'skipped': _count('skip_queue'),
        'overrides': _count('manual_overrides'),
    }
    return render_template('admin.html', counts=counts, months=dy.list_months(),
                           skip_current_month_only=dy.get_skip_current_month_only(),
                           generation_settings=dy.get_generation_settings())


@app.route('/admin/skip-policy', methods=['POST'])
def admin_skip_policy() -> str:
    enabled = request.form.get('skip_current_month_only') == 'on'
    try:
        if 'generation_mode' in request.form:
            dy.set_generation_settings(request.form['generation_mode'],
                                       int(request.form.get('prior_gap_days', '30')), enabled)
        else:
            dy.set_skip_current_month_only(enabled)
    except (ValueError, TypeError) as exc:
        flash(str(exc), 'error')
        return redirect(url_for('admin'))
    flash('배정 및 스킵 정책을 저장했습니다. 이후 생성부터 적용됩니다.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/delete-month', methods=['POST'])
def admin_delete_month() -> str:
    """지정 월의 생성 데이터를 삭제하고 배정 전 상태로 되돌린다(포인터·스킵 큐 포함).

    안전장치: 확인 문구 "삭제" + 백업 체크박스(기본 ON) + JS confirm.
    수동 배정(manual_overrides)은 월 삭제 시에도 유지된다.
    """
    year_month = request.form.get('month', '')
    if not _is_valid_ym(year_month):
        flash('연월 형식은 YYYY-MM 입니다.', 'error')
        return redirect(url_for('admin'))
    if request.form.get('confirm', '').strip() != '삭제':
        flash('취소되었습니다: 확인 문구가 정확하지 않습니다 ("삭제" 입력 필요).', 'error')
        return redirect(url_for('admin'))
    backup = dy.backup_database() if request.form.get('backup', 'off') == 'on' else None
    res = dy.delete_month(year_month)
    method_kr = {'snapshot': '스냅샷으로 생성 전 포인터 정확 복원',
                 'history': '스냅샷 없어 이력 기반 추정 롤백(재생성 시 인원 달라질 수 있음)',
                 'none': '해당 월 기록 없음(변경 없음)'}[res['pointer_restored_from']]
    msg = (f"{year_month} 삭제 완료: 이력 {res['history_deleted']}건, 스킵 {res['skip_cleared']}건 제거, "
           f"복귀대기 {res['returns_restored']}건 복구. 포인터: {method_kr}. 수동 배정은 유지됨.")
    if backup:
        msg += f' 백업: {os.path.basename(backup)}'
    flash(msg, 'success')
    return redirect(url_for('admin'))


@app.route('/admin/reset-db', methods=['POST'])
def admin_reset_db() -> str:
    """전체 DB를 빈 스키마로 초기화한다. 안전장치: 확인 문구 + 백업 옵션."""
    if request.form.get('confirm', '').strip() != '초기화':
        flash('취소되었습니다: 확인 문구가 정확하지 않습니다 ("초기화" 입력 필요).', 'error')
        return redirect(url_for('admin'))
    backup = dy.reset_database(backup=request.form.get('backup', 'off') == 'on')
    msg = '전체 DB를 초기화했습니다 (빈 스키마). "입력양식 등록"부터 다시 진행하세요.'
    if backup:
        msg += f' 백업: {os.path.basename(backup)}'
    flash(msg, 'success')
    return redirect(url_for('admin'))


if __name__ == '__main__':
    dy.migrate_v2_to_v3()
    dy.init_db()
    app.run(debug=True, port=5001)
