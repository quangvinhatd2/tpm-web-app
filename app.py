import os
import sqlite3
import re
from flask import Flask, render_template, request, redirect, url_for, session, flash
from openpyxl import load_workbook
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

        # Dữ liệu bắt đầu từ dòng 7 (theo cấu trúc file phân công bạn gửi)
        for row in range(7, ws.max_row + 1):
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

            # Người đánh giá
            if name_eval:
                uid = create_or_get_user(conn, name_eval, 'danh_gia')
                for sname in sheet_names:
                    conn.execute(
                        'INSERT INTO assignments (user_id, sheet_name, role) VALUES (?,?,?)',
                        (uid, sname, 'danh_gia')
                    )
            # Người thẩm tra
            if name_check:
                uid = create_or_get_user(conn, name_check, 'tham_tra')
                for sname in sheet_names:
                    conn.execute(
                        'INSERT INTO assignments (user_id, sheet_name, role) VALUES (?,?,?)',
                        (uid, sname, 'tham_tra')
                    )
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
        # headers: 7 dòng đầu (1-7)
        headers = [{col: ws[f'{col}{r}'].value for col in 'ABCDEF'} for r in range(1, 8)]
        rows = []
        extra = []
        for r_idx in range(10, ws.max_row + 1):
            row_data = {col: ws[f'{col}{r_idx}'].value or '' for col in 'ABCDEF'}
            if not any(str(v).strip() for v in row_data.values()):
                # phần còn lại là extra, đọc cả A-K
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
        flash('Sai tài khoản hoặc mật khẩu (Mật khẩu mặc định: 123)')
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
        evals = {(r['row_index'], r['col_letter']): r['value']
                 for r in conn.execute(
                     'SELECT row_index, col_letter, value FROM evaluations WHERE sheet_name = ?',
                     (sheet_name,)
                 )}
        comms = {r['row_index']: r['comment']
                 for r in conn.execute(
                     'SELECT row_index, comment FROM review_comments WHERE sheet_name = ?',
                     (sheet_name,)
                 )}
        s = conn.execute('SELECT * FROM suggestions WHERE sheet_name = ?', (sheet_name,)).fetchone()

    saved_header = {}  # dành cho các ô header có thể chỉnh sửa (nhiệt độ)

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
        suggestion=s['suggestion'] if s else '',
        reviewer_comment=s['reviewer_comment'] if s else '',
        reviewer_signature=s['reviewer_signature'] if s else '',
        checker_signature=s['checker_signature'] if s else '',
        enumerate=enumerate
    )

@app.route('/save', methods=['POST'])
def save():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    sn = request.form['sheet_name']
    role = request.form.get('role')
    uid = session['user_id']
    with get_db() as conn:
        if role == 'danh_gia':
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
            # Lưu suggestion và chữ ký người đánh giá (reviewer_signature)
            conn.execute(
                '''INSERT INTO suggestions (sheet_name, suggestion, reviewer_signature)
                   VALUES (?,?,?) ON CONFLICT(sheet_name) DO UPDATE SET
                   suggestion=excluded.suggestion, reviewer_signature=excluded.reviewer_signature''',
                (sn, request.form.get('suggestion', ''), request.form.get('reviewer_signature', ''))
            )
            flash('Đã lưu đánh giá thành công.')
        elif role == 'tham_tra':
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
            # Lưu nhận xét thẩm tra và chữ ký thẩm tra (checker_signature)
            conn.execute(
                '''INSERT INTO suggestions (sheet_name, reviewer_comment, checker_signature)
                   VALUES (?,?,?) ON CONFLICT(sheet_name) DO UPDATE SET
                   reviewer_comment=excluded.reviewer_comment, checker_signature=excluded.checker_signature''',
                (sn, request.form.get('reviewer_comment', ''), request.form.get('checker_signature', ''))
            )
            flash('Đã lưu ý kiến thẩm tra thành công.')
        conn.commit()
    return redirect(url_for('evaluation_form', sheet_name=sn))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    if os.path.exists(DATABASE):
        os.remove(DATABASE)
    init_db()
    app.run(debug=True)