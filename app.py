import os
import psycopg2
import re
import json
from io import BytesIO
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from unicodedata import normalize
from dotenv import load_dotenv


load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev")

# ================= FIX NEON URL =================
DATABASE_URL = os.getenv("DATABASE_URL")

if DATABASE_URL and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

# ================= ROUTES KHÔNG ĐỔI =================
@app.route('/')
def index():
    return redirect(url_for('login'))

@app.route('/export-summary')
def export_summary():
    return "Export summary"

@app.route('/export-all')
def export_all_forms():
    return "Export all"

@app.route('/reset-cycle')
def reset_cycle():
    return "Reset"

@app.route('/admin-dashboard')
def admin_dashboard():
    return "Admin dashboard"

@app.route('/sync-assignments', methods=['POST'])
def sync_assignments():
    return redirect(url_for('dashboard'))

# ================= DB =================
def get_db():
    return psycopg2.connect(
        DATABASE_URL,
        sslmode='require'   # ✅ BẮT BUỘC CHO NEON
    )

def query_all(query, params=()):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute(query, params)
            if cur.description is None:
                return []
            cols = [desc[0] for desc in cur.description]
            return [dict(zip(cols, row)) for row in cur.fetchall()]

def query_one(query, params=()):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute(query, params)
            row = cur.fetchone()
            if not row:
                return None
            cols = [desc[0] for desc in cur.description]
            return dict(zip(cols, row))

def execute(query, params=()):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute(query, params)
        conn.commit()

# ================= LOAD FORM EXCEL =================
def get_sheet_data(sheet_name):
    file_path = "forms.xlsx"

    if not os.path.exists(file_path):
        print("❌ Không tìm thấy forms.xlsx")
        return [], [], []

    wb = load_workbook(file_path, data_only=True)

    if sheet_name not in wb.sheetnames:
        print(f"❌ Không có sheet: {sheet_name}")
        return [], [], []

    ws = wb[sheet_name]

    headers = []
    rows = []
    extra = []

    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        row_data = {}
        empty = True

        for j, cell in enumerate(row):
            col_letter = chr(65 + j)
            value = "" if cell is None else str(cell).strip()

            if value != "":
                empty = False

            row_data[col_letter] = value

        if empty:
            continue

        if i <= 6:
            headers.append(row_data)
        else:
            rows.append(row_data)

    print(f"✅ Loaded sheet {sheet_name}: {len(headers)} headers, {len(rows)} rows")

    return headers, rows, extra
# ================= INIT =================
def init_db():
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE,
                password TEXT,
                fullname TEXT,
                role TEXT
            );

            CREATE TABLE IF NOT EXISTS assignments (
                id SERIAL PRIMARY KEY,
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
                checker_signature TEXT,
                locked_danh_gia INTEGER DEFAULT 0,
                locked_tham_tra INTEGER DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS archives (
                id SERIAL PRIMARY KEY,
                archive_date TEXT,
                table_name TEXT,
                row_data TEXT
            );

            CREATE TABLE IF NOT EXISTS history (
                id SERIAL PRIMARY KEY,
                sheet_name TEXT,
                role TEXT,
                user_id INTEGER,
                saved_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                snapshot TEXT
            );
            """)
        conn.commit()

# ================= LOGIN =================
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = query_one(
            'SELECT * FROM users WHERE username=%s AND password=%s',
            (request.form['username'], request.form['password'])
        )
        if user:
            session['user_id'] = user['id']
            session['fullname'] = user['fullname']
            session['role'] = user['role']
            return redirect(url_for('dashboard'))
        flash('Sai tài khoản hoặc mật khẩu')
    return render_template('login.html')

# ================= DASHBOARD =================
@app.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    assigns = query_all("""
        SELECT a.sheet_name, a.role,
        COALESCE(s.locked_danh_gia,0) as locked_danh_gia,
        COALESCE(s.locked_tham_tra,0) as locked_tham_tra
        FROM assignments a
        LEFT JOIN suggestions s ON a.sheet_name=s.sheet_name
        WHERE a.user_id=%s
    """, (session['user_id'],))

    eval_status = {}

    for ass in assigns:
        result = query_one("""
            SELECT 1 FROM evaluations
            WHERE sheet_name=%s AND user_id=%s
            LIMIT 1
        """, (ass['sheet_name'], session['user_id']))

        eval_status[ass['sheet_name']] = True if result else False

    return render_template(
        'dashboard.html',
        assignments=assigns,
        eval_status=eval_status
    )

# ================= SAVE =================
@app.route('/save', methods=['POST'])
def save():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    sn = request.form.get('sheet_name')
    if not sn:
        flash("Thiếu sheet_name")
        return redirect(url_for('dashboard'))

    uid = session['user_id']

    for k, v in request.form.items():
        if k.startswith('eval_'):
            parts = k.split('_')
            if len(parts) == 3:
                row = int(parts[1])
                col = parts[2]

                execute("""
                INSERT INTO evaluations (user_id, sheet_name, row_index, col_letter, value)
                VALUES (%s,%s,%s,%s,%s)
                ON CONFLICT (user_id, sheet_name, row_index, col_letter)
                DO UPDATE SET value=EXCLUDED.value
                """, (uid, sn, row, col, v))

    flash('Đã lưu')
    return redirect(url_for('evaluation_form', sheet_name=sn))

# ================= HISTORY =================
@app.route('/history')
def history():
    rows = query_all("""
        SELECT h.*, u.fullname
        FROM history h
        JOIN users u ON h.user_id=u.id
        ORDER BY h.saved_at DESC
    """)
    return render_template('history.html', history=rows)

# ================= LOGOUT (FIX CỨNG) =================
@app.route('/logout', endpoint='logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ================= EVALUATION =================
@app.route('/evaluation/<sheet_name>', methods=['GET', 'POST'])
def evaluation_form(sheet_name):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    uid = session['user_id']

    if request.method == 'POST':
        for key, value in request.form.items():
            if key.startswith('eval_'):
                parts = key.split('_')
                if len(parts) == 3:
                    row = int(parts[1])
                    col = parts[2]

                    execute("""
                    INSERT INTO evaluations (user_id, sheet_name, row_index, col_letter, value)
                    VALUES (%s,%s,%s,%s,%s)
                    ON CONFLICT (user_id, sheet_name, row_index, col_letter)
                    DO UPDATE SET value=EXCLUDED.value
                    """, (uid, sheet_name, row, col, value))

        flash('Đã lưu')
        return redirect(url_for('evaluation_form', sheet_name=sheet_name))

    data = query_all("""
        SELECT row_index, col_letter, value
        FROM evaluations
        WHERE user_id=%s AND sheet_name=%s
    """, (uid, sheet_name))

    saved = {(r['row_index'], r['col_letter']): r['value'] for r in data} if data else {}

    headers, rows, extra = get_sheet_data(sheet_name)

    if not headers:
        return f"❌ Sheet '{sheet_name}' không có dữ liệu hoặc sai tên", 404

    return render_template(
    'evaluation_form.html',
    sheet_name=sheet_name,
    role=session.get('role', 'danh_gia'),
    headers=headers,
    rows=rows,
    extra=extra,
    saved=saved,
    saved_header={},
    saved_comments={},
    reviewer_signature='',
    checker_signature='',
    locked_danh_gia=0,
    locked_tham_tra=0,
    enumerate=enumerate        # ← thêm dòng này
)

# ================= RUN =================
if __name__ == '__main__':
    
    init_db()
    print(app.url_map)  # DEBUG ROUTES
    app.run(debug=True, host='0.0.0.0', port=5000)