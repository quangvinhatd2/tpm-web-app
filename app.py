import os
import sqlite3
import re
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
from unicodedata import normalize

app = Flask(__name__)
app.secret_key = 'tpm-secure-key-2024'
DATABASE = os.path.join(app.instance_path, 'app.db')

if not os.path.exists(app.instance_path):
    os.makedirs(app.instance_path)

FORMS_FILE = 'forms.xlsx'
PHAN_GIAO_FILE = 'phan_giao.xlsx'

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn

def unsigned_user(text):
    if not text:
        return None
    text = normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'[^a-zA-Z0-9]', '', text).lower()

def build_sheet_mapping():
    wb = load_workbook(FORMS_FILE, data_only=True)
    mapping = {}
    for sheet_name in wb.sheetnames:
        if sheet_name.startswith('BM'):
            code_part = sheet_name[2:]
            if code_part.isdigit():
                num = int(code_part)
                mapping[f'BM.P4.15.{num:02d}'] = sheet_name
            elif '_' in code_part:
                base, pha = code_part.split('_')
                num = int(base)
                mapping[f'BM.P4.15.{num:02d}_{pha}'] = sheet_name
    wb.close()
    return mapping

def build_reverse_mapping():
    wb = load_workbook(FORMS_FILE, data_only=True)
    rev_map = {}
    for sheet_name in wb.sheetnames:
        if sheet_name.startswith('BM'):
            code_part = sheet_name[2:]
            if code_part.isdigit():
                num = int(code_part)
                rev_map[sheet_name] = f'BM.P4.15.{num:02d}'
            elif '_' in code_part:
                base, pha = code_part.split('_')
                num = int(base)
                rev_map[sheet_name] = f'BM.P4.15.{num:02d}_{pha}'
    wb.close()
    return rev_map

def init_db():
    if not os.path.exists(PHAN_GIAO_FILE):
        print("Không tìm thấy file phan_giao.xlsx")
        return

    mapping = build_sheet_mapping()

    with get_db() as conn:
        conn.executescript('''
            DROP TABLE IF EXISTS users;
            DROP TABLE IF EXISTS assignments;
            CREATE TABLE users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE,
                password TEXT,
                fullname TEXT,
                role TEXT
            );
            CREATE TABLE assignments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                sheet_name TEXT,
                role TEXT
            );
            CREATE TABLE IF NOT EXISTS evaluations (
                user_id INTEGER,
                sheet_name TEXT,
                row_index INTEGER,
                col_letter TEXT,
                value TEXT,
                PRIMARY KEY (user_id, sheet_name, row_index, col_letter)
            );
            CREATE TABLE IF NOT EXISTS review_comments (
                reviewer_id INTEGER,
                sheet_name TEXT,
                row_index INTEGER,
                comment TEXT,
                PRIMARY KEY (reviewer_id, sheet_name, row_index)
            );
            CREATE TABLE IF NOT EXISTS suggestions (
                sheet_name TEXT PRIMARY KEY,
                suggestion TEXT,
                reviewer_comment TEXT,
                reviewer_signature TEXT,
                checker_signature TEXT
            );
        ''')

        wb = load_workbook(PHAN_GIAO_FILE, data_only=True)
        ws = wb.active

        for row in range(8, ws.max_row + 1):
            ma_bieu_mau = str(ws[f'C{row}'].value or '').strip()
            name_eval = str(ws[f'E{row}'].value or '').strip()
            name_check = str(ws[f'F{row}'].value or '').strip()

            if not ma_bieu_mau or (not name_eval and not name_check):
                continue

            if ma_bieu_mau in ['BM.P4.15.18', 'BM.P4.15.19']:
                base_num = ma_bieu_mau.split('.')[-1]
                sheet_names = [f'BM{base_num}_a', f'BM{base_num}_b', f'BM{base_num}_c']
            else:
                sheet_name = mapping.get(ma_bieu_mau)
                if not sheet_name:
                    print(f"Không tìm thấy sheet cho mã {ma_bieu_mau}")
                    continue
                sheet_names = [sheet_name]

            if name_eval:
                uid = create_or_get_user(conn, name_eval, 'danh_gia')
                for sname in sheet_names:
                    conn.execute(
                        'INSERT INTO assignments (user_id, sheet_name, role) VALUES (?,?,?)',
                        (uid, sname, 'danh_gia')
                    )
            if name_check:
                uid = create_or_get_user(conn, name_check, 'tham_tra')
                for sname in sheet_names:
                    conn.execute(
                        'INSERT INTO assignments (user_id, sheet_name, role) VALUES (?,?,?)',
                        (uid, sname, 'tham_tra')
                    )

        # Tạo tài khoản admin nếu chưa tồn tại
        admin_username = 'admin'
        admin_exists = conn.execute('SELECT id FROM users WHERE username = ?', (admin_username,)).fetchone()
        if not admin_exists:
            conn.execute(
                'INSERT INTO users (username, password, fullname, role) VALUES (?,?,?,?)',
                (admin_username, 'admin123', 'Quản trị viên', 'admin')
            )
            print("Đã tạo tài khoản admin (user: admin, pass: admin123)")

        conn.commit()
    print("--- Đã nạp dữ liệu phân công thành công ---")

def create_or_get_user(conn, fullname, role):
    username = unsigned_user(fullname)
    user = conn.execute('SELECT id FROM users WHERE username = ?', (username,)).fetchone()
    if user:
        return user['id']
    cursor = conn.execute(
        'INSERT INTO users (username, password, fullname, role) VALUES (?,?,?,?)',
        (username, '123', fullname, role)
    )
    return cursor.lastrowid

def get_sheet_data(sheet_name):
    try:
        wb = load_workbook(FORMS_FILE, data_only=True)
        if sheet_name not in wb.sheetnames:
            return None, None, None
        ws = wb[sheet_name]
        headers = [{col: ws[f'{col}{r}'].value for col in 'ABCDEF'} for r in range(1, 8)]
        rows = []
        extra = []
        for r_idx in range(10, ws.max_row + 1):
            row_data = {col: ws[f'{col}{r_idx}'].value or '' for col in 'ABCDEF'}
            if not any(str(v).strip() for v in row_data.values()):
                for r in range(r_idx, ws.max_row + 1):
                    e_row = {col: ws[f'{col}{r}'].value or '' for col in 'ABCDEFGHIJK'}
                    if any(str(v).strip() for v in e_row.values()):
                        extra.append(e_row)
                break
            rows.append(row_data)
        return headers, rows, extra
    except Exception as e:
        print(f"Lỗi đọc sheet {sheet_name}: {e}")
        return None, None, None

# -------------------- ROUTES --------------------
@app.route('/')
def index():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        with get_db() as conn:
            user = conn.execute(
                'SELECT * FROM users WHERE username = ? AND password = ?',
                (username, password)
            ).fetchone()
        if user:
            session['user_id'] = user['id']
            session['fullname'] = user['fullname']
            session['role'] = user['role']
            return redirect(url_for('dashboard'))
        flash('Sai tài khoản hoặc mật khẩu (Mật khẩu mặc định: 123 cho user, admin: admin123)')
    return render_template('login.html')

@app.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    with get_db() as conn:
        assigns = conn.execute(
            'SELECT sheet_name, role FROM assignments WHERE user_id = ?',
            (session['user_id'],)
        ).fetchall()
    return render_template('dashboard.html', assignments=assigns)

@app.route('/form/<sheet_name>')
def evaluation_form(sheet_name):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    with get_db() as conn:
        assign = conn.execute(
            'SELECT role FROM assignments WHERE user_id = ? AND sheet_name = ?',
            (session['user_id'], sheet_name)
        ).fetchone()
    if not assign:
        return "Bạn không có quyền truy cập biểu mẫu này", 403

    headers, rows, extra = get_sheet_data(sheet_name)
    if not headers:
        return f"Không tìm thấy sheet {sheet_name} trong file forms.xlsx", 404

    with get_db() as conn:
        # Lấy dữ liệu đánh giá và header đã lưu
        db_rows = conn.execute(
                     'SELECT row_index, col_letter, value FROM evaluations WHERE sheet_name = ?',
                     (sheet_name,)
                 ).fetchall()
        
        evals = {(r['row_index'], r['col_letter']): r['value'] for r in db_rows if r['row_index'] >= 10}
        saved_header = {(r['row_index'], r['col_letter']): r['value'] for r in db_rows if r['row_index'] < 10}
        
        comms = {r['row_index']: r['comment']
                 for r in conn.execute(
                     'SELECT row_index, comment FROM review_comments WHERE sheet_name = ?',
                     (sheet_name,)
                 ).fetchall()}
        s = conn.execute('SELECT * FROM suggestions WHERE sheet_name = ?', (sheet_name,)).fetchone()

    # Xử lý giá trị None từ DB để tránh hiển thị chữ 'None' trên giao diện
    suggestion = (s['suggestion'] if s else '') or ''
    reviewer_comment = (s['reviewer_comment'] if s else '') or ''
    reviewer_signature = (s['reviewer_signature'] if s else '') or ''
    checker_signature = (s['checker_signature'] if s else '') or ''

    return render_template(
        'evaluation_form.html',
        sheet_name=sheet_name,
        role=assign['role'],
        headers=headers,
        rows=rows,
        extra=extra,
        saved=evals,
        saved_comments=comms,
        saved_header=saved_header,
        suggestion=suggestion,
        reviewer_comment=reviewer_comment,
        reviewer_signature=reviewer_signature,
        checker_signature=checker_signature,
        enumerate=enumerate
    )

@app.route('/save', methods=['POST'])
def save():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    sn = request.form['sheet_name']
    role = request.form.get('role')
    uid = session['user_id']

    # Hàm chuyển tên cột sang tên hiển thị
    def col_name(col):
        return {
            'H': 'Mô tả',
            'I': 'Đơn vị thực hiện',
            'J': 'Thời gian',
            'K': 'Giải pháp'
        }.get(col, col)

    # Kiểm tra và lưu dữ liệu header (nhiệt độ)
    with get_db() as conn:
        if role == 'danh_gia':
            cycle_val = request.form.get('header_6_E', '').strip()
            if not cycle_val:
                flash('Vui lòng nhập "Nhiệt độ môi trường - Kiểm tra" (ô đầu tiên).')
                return redirect(url_for('evaluation_form', sheet_name=sn))
            conn.execute(
                'INSERT OR REPLACE INTO evaluations (user_id, sheet_name, row_index, col_letter, value) VALUES (?,?,?,?,?)',
                (uid, sn, 6, 'E', cycle_val)
            )
        elif role == 'tham_tra':
            cycle_val = request.form.get('header_6_F', '').strip()
            if not cycle_val:
                flash('Vui lòng nhập "Nhiệt độ môi trường - Thẩm tra" (ô thứ hai).')
                return redirect(url_for('evaluation_form', sheet_name=sn))
            conn.execute(
                'INSERT OR REPLACE INTO evaluations (user_id, sheet_name, row_index, col_letter, value) VALUES (?,?,?,?,?)',
                (uid, sn, 6, 'F', cycle_val)
            )
        conn.commit()

    # Thu thập các ô eval
    eval_items = {}
    for key, value in request.form.items():
        if key.startswith('eval_'):
            parts = key.split('_')
            if len(parts) == 3:
                try:
                    row = int(parts[1])
                    col = parts[2]
                except:
                    continue
                if row not in eval_items:
                    eval_items[row] = {}
                eval_items[row][col] = value

    if role == 'danh_gia':
        # Kiểm tra từng dòng nội dung
        for row, cols in eval_items.items():
            if 'G' not in cols or not cols['G'].strip():
                flash(f'Dòng {row}: chưa chọn kết quả (cột G).')
                return redirect(url_for('evaluation_form', sheet_name=sn))
            if cols.get('G') == 'K':
                missing = []
                for col in ['H','I','J','K']:
                    if col not in cols or not cols[col].strip():
                        missing.append(col_name(col))
                if missing:
                    flash(f'Dòng {row} (kết quả K) còn thiếu các cột: {", ".join(missing)}.')
                    return redirect(url_for('evaluation_form', sheet_name=sn))

        # Kiểm tra kiến nghị của người đánh giá
        reviewer_sig = request.form.get('reviewer_signature', '').strip()
        if not reviewer_sig:
            flash('Vui lòng nhập nội dung tại ô "Người đánh giá" (ký xác nhận).')
            return redirect(url_for('evaluation_form', sheet_name=sn))

        # Lưu dữ liệu đánh giá chi tiết
        with get_db() as conn:
            for k, v in request.form.items():
                if k.startswith('eval_'):
                    parts = k.split('_')
                    if len(parts) == 3:
                        try:
                            row = int(parts[1])
                            col = parts[2]
                        except:
                            continue
                        conn.execute(
                            'INSERT OR REPLACE INTO evaluations (user_id, sheet_name, row_index, col_letter, value) VALUES (?,?,?,?,?)',
                            (uid, sn, row, col, v)
                        )
            conn.execute(
                '''INSERT INTO suggestions (sheet_name, suggestion, reviewer_signature)
                   VALUES (?,?,?) ON CONFLICT(sheet_name) DO UPDATE SET
                   suggestion=excluded.suggestion, reviewer_signature=excluded.reviewer_signature''',
                (sn, request.form.get('suggestion', ''), reviewer_sig)
            )
            conn.commit()
        flash('Đã lưu đánh giá thành công.')

    elif role == 'tham_tra':
        # Kiểm tra ý kiến thẩm tra từng dòng
        comment_items = {}
        for key, value in request.form.items():
            if key.startswith('comment_'):
                parts = key.split('_')
                if len(parts) == 2:
                    try:
                        row = int(parts[1])
                    except:
                        continue
                    comment_items[row] = value
        
        with get_db() as conn:
            rows_to_check = conn.execute('SELECT DISTINCT row_index FROM evaluations WHERE sheet_name=? AND col_letter="G"', (sn,)).fetchall()
        
        for r in rows_to_check:
            row = r['row_index']
            if row not in comment_items or not comment_items[row].strip():
                flash(f'Dòng {row}: chưa nhập ý kiến thẩm tra.')
                return redirect(url_for('evaluation_form', sheet_name=sn))

        # Kiểm tra chữ ký thẩm tra
        checker_sig = request.form.get('checker_signature', '').strip()
        if not checker_sig:
            flash('Vui lòng nhập nội dung tại ô "Người thẩm tra" (ký xác nhận).')
            return redirect(url_for('evaluation_form', sheet_name=sn))

        # Lưu dữ liệu thẩm tra
        with get_db() as conn:
            for k, v in request.form.items():
                if k.startswith('comment_'):
                    parts = k.split('_')
                    if len(parts) == 2:
                        try:
                            row = int(parts[1])
                        except:
                            continue
                        conn.execute(
                            'INSERT OR REPLACE INTO review_comments (reviewer_id, sheet_name, row_index, comment) VALUES (?,?,?,?)',
                            (uid, sn, row, v)
                        )
            conn.execute(
                '''INSERT INTO suggestions (sheet_name, reviewer_comment, checker_signature)
                   VALUES (?,?,?) ON CONFLICT(sheet_name) DO UPDATE SET
                   reviewer_comment=excluded.reviewer_comment, checker_signature=excluded.checker_signature''',
                (sn, request.form.get('reviewer_comment', ''), checker_sig)
            )
            conn.commit()
        flash('Đã lưu ý kiến thẩm tra thành công.')

    return redirect(url_for('evaluation_form', sheet_name=sn))

@app.route('/export_summary')
def export_summary():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền truy cập chức năng này.')
        return redirect(url_for('dashboard'))

    with get_db() as conn:
        rows = conn.execute('''
            SELECT e.sheet_name, e.row_index,
                   e.value as result,
                   (SELECT value FROM evaluations e2 
                    WHERE e2.sheet_name = e.sheet_name 
                      AND e2.row_index = e.row_index 
                      AND e2.col_letter = 'H') as description,
                   (SELECT comment FROM review_comments rc 
                    WHERE rc.sheet_name = e.sheet_name 
                      AND rc.row_index = e.row_index) as reviewer_comment
            FROM evaluations e
            WHERE e.col_letter = 'G' AND e.value = 'K'
            ORDER BY e.sheet_name, e.row_index
        ''').fetchall()

        suggestions = conn.execute('''
            SELECT sheet_name, reviewer_signature, checker_signature
            FROM suggestions
            WHERE sheet_name IN (SELECT DISTINCT sheet_name 
                                 FROM evaluations 
                                 WHERE col_letter = 'G' AND value = 'K')
        ''').fetchall()
        sug_dict = {s['sheet_name']: (s['reviewer_signature'], s['checker_signature']) for s in suggestions}

    if not rows:
        flash('Không có khiếm khuyết nào để xuất báo cáo.')
        return redirect(url_for('dashboard'))

    rev_map = build_reverse_mapping()

    wb = Workbook()
    ws = wb.active
    ws.title = "Tổng hợp KKTB"

    ws.merge_cells('A1:D1')
    ws['A1'] = "TỔNG HỢP KHIẾM KHUYẾT THIẾT BỊ TPM"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    headers = ["STT", "Biểu mẫu", "Nội dung khiếm khuyết (Mô tả của ĐG)", "Mô tả của Thẩm tra"]
    ws.append(headers)
    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    stt = 1
    for row in rows:
        sheet_code = rev_map.get(row['sheet_name'], row['sheet_name'])
        ws.append([
            stt,
            sheet_code,
            row['description'] or '',
            row['reviewer_comment'] or ''
        ])
        stt += 1

    for col in ws.columns:
        max_length = 0
        col_letter = None
        for cell in col:
            if hasattr(cell, 'column_letter'):
                col_letter = cell.column_letter
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
        if col_letter:
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[col_letter].width = adjusted_width

    if sug_dict:
        ws.append([])
        ws.append(["KIẾN NGHỊ VÀ Ý KIẾN THẨM TRA"])
        ws.merge_cells(f'A{ws.max_row}:D{ws.max_row}')
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        ws.cell(row=ws.max_row, column=1).alignment = Alignment(horizontal='center')

        for sheet_code, (reviewer_sug, checker_comm) in sug_dict.items():
            ws.append([sheet_code, "Kiến nghị của người đánh giá:", reviewer_sug or '', ""])
            ws.append([sheet_code, "Ý kiến của người thẩm tra:", checker_comm or '', ""])
            for row in range(ws.max_row-1, ws.max_row+1):
                for cell in ws[row]:
                    cell.alignment = Alignment(horizontal='left', wrap_text=True)
        for col in ws.columns:
            max_length = 0
            col_letter = None
            for cell in col:
                if hasattr(cell, 'column_letter'):
                    col_letter = cell.column_letter
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
            if col_letter:
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[col_letter].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(output,
                     download_name='TonghopKKTB TPM.xlsx',
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    # init_db()  # Chỉ mở lại nếu muốn reset database từ file phan_giao.xlsx
    app.run(debug=True, host='0.0.0.0')