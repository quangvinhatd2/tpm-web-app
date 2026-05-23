import os
import re
import json
import threading
from io import BytesIO
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from unicodedata import normalize
import psycopg2
from psycopg2.pool import SimpleConnectionPool
from psycopg2.extras import RealDictCursor
from contextlib import contextmanager
from dotenv import load_dotenv
import time

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'fallback-local-key')

DATABASE_URL = os.environ.get('DATABASE_URL')
if not DATABASE_URL:
    raise Exception("DATABASE_URL environment variable not set")

FORMS_FILE = 'forms.xlsx'

MASTER_WB = None
MASTER_WB_MTIME = None

PHAN_GIAO_FILE = 'phan_giao.xlsx'

_sheet_cache = {}


def get_master_workbook():

    global MASTER_WB
    global MASTER_WB_MTIME
    global _sheet_cache

    current_mtime = os.path.getmtime(FORMS_FILE)

    # File chưa đổi → dùng cache
    if (
        MASTER_WB is not None
        and MASTER_WB_MTIME == current_mtime
    ):
        return MASTER_WB

    # File đổi → reload
    if MASTER_WB:
        MASTER_WB.close()

    MASTER_WB = load_workbook(
        FORMS_FILE,
        read_only=True,
        data_only=True
    )

    MASTER_WB_MTIME = current_mtime

    # XÓA CACHE CŨ
    _sheet_cache.clear()

    print("✅ Workbook cache reloaded")

    return MASTER_WB

# ================= KẾT NỐI MỚI (không pool, có retry) =================
# ================= DATABASE CONNECTION POOL =================

DB_POOL_MIN = 1
DB_POOL_MAX = 20

max_retries = 3

for i in range(max_retries):

    try:

        db_pool = SimpleConnectionPool(
            minconn=DB_POOL_MIN,
            maxconn=DB_POOL_MAX,
            dsn=DATABASE_URL,
            sslmode='require',
            cursor_factory=RealDictCursor
        )

        print(f"✅ Database pool initialized ({DB_POOL_MIN}-{DB_POOL_MAX})")

        break

    except Exception as e:

        if i == max_retries - 1:
            raise

        print(f"❌ Lỗi tạo connection pool: {e}")

        time.sleep(1 * (2 ** i))


@contextmanager
def get_db_connection():

    conn = None

    try:

        conn = db_pool.getconn()

        yield conn

        conn.commit()

    except Exception:

        if conn:
            conn.rollback()

        raise

    finally:

        if conn:
            db_pool.putconn(conn)

# -------------------- Hàm xử lý Excel an toàn --------------------
def safe_load_workbook(filepath, read_only=False):
    """Load workbook với xử lý lỗi và read_only cho file lớn"""
    try:
        if read_only:
            return load_workbook(filepath, read_only=True, data_only=True)
        else:
            return load_workbook(filepath, data_only=True)
    except Exception as e:
        print(f"Lỗi đọc file {filepath}: {e}")
        return None

def build_sheet_mapping():
    wb = get_master_workbook()
    if not wb:
        return {}
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
    return mapping

def build_reverse_mapping():
    wb = get_master_workbook()
    if not wb:
        return {}
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
    return rev_map

def get_sheet_data(sheet_name):
    # Trả về từ cache nếu đã đọc rồi
    if sheet_name in _sheet_cache:
        return _sheet_cache[sheet_name]
    try:
        wb = get_master_workbook()
        if not wb or sheet_name not in wb.sheetnames:
            return None, None, None
        ws = wb[sheet_name]
        headers = [{col: ws[f'{col}{r}'].value for col in 'ABCDEF'} for r in range(1, 8)]
        rows = []
        extra = []
        max_row = min(ws.max_row, 500)
        for r_idx in range(10, max_row + 1):
            row_data = {col: ws[f'{col}{r_idx}'].value or '' for col in 'ABCDEF'}
            if not any(str(v).strip() for v in row_data.values()):
                break
            rows.append(row_data)
        # Lưu vào cache
        _sheet_cache[sheet_name] = (headers, rows, extra)
        return headers, rows, extra
    except Exception as e:
        print(f"Lỗi đọc sheet {sheet_name}: {e}")
        return None, None, None

# -------------------- Các hàm xử lý database --------------------
def unsigned_user(text):
    if not text:
        return None
    text = text.strip()
    replacements = {
        'Đ': 'D', 'đ': 'd',
        'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A',
        'à': 'a', 'á': 'a', 'â': 'a', 'ã': 'a',
        'È': 'E', 'É': 'E', 'Ê': 'E',
        'è': 'e', 'é': 'e', 'ê': 'e',
        'Ì': 'I', 'Í': 'I', 'ì': 'i', 'í': 'i',
        'Ò': 'O', 'Ó': 'O', 'Ô': 'O', 'Õ': 'O',
        'ò': 'o', 'ó': 'o', 'ô': 'o', 'õ': 'o',
        'Ù': 'U', 'Ú': 'U', 'ù': 'u', 'ú': 'u',
        'Ý': 'Y', 'ý': 'y',
        'Ă': 'A', 'ă': 'a', 'Ắ': 'A', 'ắ': 'a', 'Ặ': 'A', 'ặ': 'a',
        'Ằ': 'A', 'ằ': 'a', 'Ẳ': 'A', 'ẳ': 'a', 'Ẵ': 'A', 'ẵ': 'a',
        'Ấ': 'A', 'ấ': 'a', 'Ầ': 'A', 'ầ': 'a', 'Ẩ': 'A', 'ẩ': 'a',
        'Ẫ': 'A', 'ẫ': 'a', 'Ậ': 'A', 'ậ': 'a',
        'Ơ': 'O', 'ơ': 'o', 'Ớ': 'O', 'ớ': 'o', 'Ờ': 'O', 'ờ': 'o',
        'Ở': 'O', 'ở': 'o', 'Ỡ': 'O', 'ỡ': 'o', 'Ợ': 'O', 'ợ': 'o',
        'Ố': 'O', 'ố': 'o', 'Ồ': 'O', 'ồ': 'o', 'Ổ': 'O', 'ổ': 'o',
        'Ỗ': 'O', 'ỗ': 'o', 'Ộ': 'O', 'ộ': 'o',
        'Ư': 'U', 'ư': 'u', 'Ứ': 'U', 'ứ': 'u', 'Ừ': 'U', 'ừ': 'u',
        'Ử': 'U', 'ử': 'u', 'Ữ': 'U', 'ữ': 'u', 'Ự': 'U', 'ự': 'u',
        'Ế': 'E', 'ế': 'e', 'Ề': 'E', 'ề': 'e', 'Ể': 'E', 'ể': 'e',
        'Ễ': 'E', 'ễ': 'e', 'Ệ': 'E', 'ệ': 'e',
        'Ỉ': 'I', 'ỉ': 'i', 'Ị': 'I', 'ị': 'i',
        'Ỳ': 'Y', 'ỳ': 'y', 'Ỷ': 'Y', 'ỷ': 'y',
        'Ỹ': 'Y', 'ỹ': 'y', 'Ỵ': 'Y', 'ỵ': 'y',
        'Ả': 'A', 'ả': 'a', 'Ạ': 'A', 'ạ': 'a',
        'Ẻ': 'E', 'ẻ': 'e', 'Ẽ': 'E', 'ẽ': 'e', 'Ẹ': 'E', 'ẹ': 'e',
        'Ỏ': 'O', 'ỏ': 'o', 'Ọ': 'O', 'ọ': 'o',
        'Ủ': 'U', 'ủ': 'u', 'Ụ': 'U', 'ụ': 'u',
    }
    result = ''
    for ch in text:
        if ch in replacements:
            result += replacements[ch]
        else:
            result += normalize('NFKD', ch).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'[^a-zA-Z0-9]', '', result).lower()

def init_db():
    if not os.path.exists(PHAN_GIAO_FILE):
        print("Không tìm thấy file phan_giao.xlsx")
        return
    mapping = build_sheet_mapping()
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("DROP TABLE IF EXISTS users CASCADE")
        cur.execute("DROP TABLE IF EXISTS assignments CASCADE")
        cur.execute("""
            CREATE TABLE users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE,
                password TEXT,
                fullname TEXT,
                role TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE assignments (
                id SERIAL PRIMARY KEY,
                user_id INTEGER,
                sheet_name TEXT,
                role TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS evaluations (
                user_id INTEGER,
                sheet_name TEXT,
                row_index INTEGER,
                col_letter TEXT,
                value TEXT,
                PRIMARY KEY (user_id, sheet_name, row_index, col_letter)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS review_comments (
                reviewer_id INTEGER,
                sheet_name TEXT,
                row_index INTEGER,
                comment TEXT,
                PRIMARY KEY (reviewer_id, sheet_name, row_index)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS suggestions (
                sheet_name TEXT PRIMARY KEY,
                suggestion TEXT,
                reviewer_comment TEXT,
                reviewer_signature TEXT,
                checker_signature TEXT,
                locked_danh_gia INTEGER DEFAULT 0,
                locked_tham_tra INTEGER DEFAULT 0
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS archives (
                id SERIAL PRIMARY KEY,
                archive_date TEXT NOT NULL,
                table_name TEXT NOT NULL,
                row_data TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS history (
                id SERIAL PRIMARY KEY,
                sheet_name TEXT NOT NULL,
                role TEXT NOT NULL,
                user_id INTEGER NOT NULL,
                saved_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                snapshot TEXT NOT NULL
            )
        """)
        conn.commit()

        wb = safe_load_workbook(PHAN_GIAO_FILE)
        if not wb:
            print("Không thể đọc file phan_giao.xlsx")
            return
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
                    continue
                sheet_names = [sheet_name]
            if name_eval:
                uid = create_or_get_user(conn, name_eval, 'danh_gia')
                for sname in sheet_names:
                    cur.execute("INSERT INTO assignments (user_id, sheet_name, role) VALUES (%s, %s, %s)", (uid, sname, 'danh_gia'))
            if name_check:
                uid = create_or_get_user(conn, name_check, 'tham_tra')
                for sname in sheet_names:
                    cur.execute("INSERT INTO assignments (user_id, sheet_name, role) VALUES (%s, %s, %s)", (uid, sname, 'tham_tra'))
        wb.close()
        cur.execute("SELECT id FROM users WHERE username = %s", ('admin',))
        if not cur.fetchone():
            cur.execute("INSERT INTO users (username, password, fullname, role) VALUES (%s, %s, %s, %s)", ('admin', 'admin123', 'Quản trị viên', 'admin'))
        conn.commit()
    print("--- Đã nạp dữ liệu phân công thành công ---")

def create_or_get_user(conn, fullname, role):
    username = unsigned_user(fullname)
    cur = conn.cursor()
    cur.execute("SELECT id FROM users WHERE username = %s", (username,))
    user = cur.fetchone()
    if user:
        cur.execute("UPDATE users SET fullname = %s WHERE username = %s", (fullname, username))
        conn.commit()
        return user['id']
    cur.execute("INSERT INTO users (username, password, fullname, role) VALUES (%s, %s, %s, %s) RETURNING id", (username, '123', fullname, role))
    user_id = cur.fetchone()['id']
    conn.commit()
    return user_id

def is_evaluation_complete(sheet_name):
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) as cnt FROM evaluations WHERE sheet_name = %s AND col_letter = 'G' AND value != '' AND value IS NOT NULL", (sheet_name,))
        has_results = cur.fetchone()['cnt']
        cur.execute("SELECT reviewer_signature FROM suggestions WHERE sheet_name = %s", (sheet_name,))
        has_signature = cur.fetchone()
    sig_ok = (has_signature is not None and has_signature['reviewer_signature'] and has_signature['reviewer_signature'].strip() != '')
    return has_results > 0 and sig_ok

def ensure_archive_table():
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS archives (
                id SERIAL PRIMARY KEY,
                archive_date TEXT NOT NULL,
                table_name TEXT NOT NULL,
                row_data TEXT NOT NULL
            )
        """)
        conn.commit()

def archive_current_data():
    ensure_archive_table()
    archive_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM evaluations")
        for row in cur.fetchall():
            cur.execute("INSERT INTO archives (archive_date, table_name, row_data) VALUES (%s, %s, %s)",
                        (archive_date, 'evaluations', json.dumps(dict(row), ensure_ascii=False)))

        cur.execute("SELECT * FROM review_comments")
        for row in cur.fetchall():
            cur.execute("INSERT INTO archives (archive_date, table_name, row_data) VALUES (%s, %s, %s)",
                        (archive_date, 'review_comments', json.dumps(dict(row), ensure_ascii=False)))

        cur.execute("SELECT * FROM suggestions")
        for row in cur.fetchall():
            cur.execute("INSERT INTO archives (archive_date, table_name, row_data) VALUES (%s, %s, %s)",
                        (archive_date, 'suggestions', json.dumps(dict(row), ensure_ascii=False)))

        # Đặt ra ngoài toàn bộ vòng lặp — chạy đúng 1 lần sau khi insert xong hết
        cur.execute("""
            DELETE FROM archives
            WHERE archive_date::timestamp < NOW() - INTERVAL '12 months'
        """)
        conn.commit()
    print(f"Đã sao lưu dữ liệu vào archive ngày {archive_date}")

def snapshot_all_before_reset():
    """
    Trước khi reset tháng mới: tạo history snapshot cho TẤT CẢ sheet đã có dữ liệu
    trong tháng hiện tại mà chưa có history entry.
    """
    try:
        with get_db_connection() as conn:
            cur = conn.cursor()
            current_month = datetime.now().strftime('%Y-%m')
            cur.execute("""
                SELECT DISTINCT sheet_name FROM evaluations
                WHERE col_letter = 'G' AND value IS NOT NULL AND value != ''
            """)
            sheets_with_evals = {r['sheet_name'] for r in cur.fetchall()}

            cur.execute("""
    SELECT DISTINCT sheet_name, role FROM history
    WHERE TO_CHAR(saved_at, 'YYYY-MM') = %s
""", (current_month,))
            already_saved = {(r['sheet_name'], r['role']) for r in cur.fetchall()}

            count = 0
            for sn in sheets_with_evals:
                cur.execute("""
                    SELECT row_index, col_letter, value FROM evaluations
                    WHERE sheet_name = %s
                """, (sn,))
                evals_rows = cur.fetchall()

                cur.execute("""
                    SELECT row_index, comment FROM review_comments
                    WHERE sheet_name = %s
                """, (sn,))
                comms_rows = cur.fetchall()

                cur.execute("SELECT * FROM suggestions WHERE sheet_name = %s", (sn,))
                sugg = cur.fetchone()

                snapshot_data = {
                    'evals': [{'row': r['row_index'], 'col': r['col_letter'], 'value': r['value']} for r in evals_rows],
                    'comments': [{'row': r['row_index'], 'comment': r['comment']} for r in comms_rows],
                    'suggestions': dict(sugg) if sugg else {}
                }
                snapshot_json = json.dumps(snapshot_data, ensure_ascii=False, default=str)

                cur.execute("""
                    SELECT a.user_id FROM assignments a
                    WHERE a.sheet_name = %s AND a.role = 'danh_gia'
                    LIMIT 1
                """, (sn,))
                dg_row = cur.fetchone()

                cur.execute("""
                    SELECT a.user_id FROM assignments a
                    WHERE a.sheet_name = %s AND a.role = 'tham_tra'
                    LIMIT 1
                """, (sn,))
                tt_row = cur.fetchone()

                if (sn, 'danh_gia') not in already_saved and dg_row:
                    cur.execute("""
                        INSERT INTO history (sheet_name, role, user_id, snapshot)
                        VALUES (%s, %s, %s, %s)
                    """, (sn, 'danh_gia', dg_row['user_id'], snapshot_json))
                    count += 1

                if (sn, 'tham_tra') not in already_saved and tt_row and comms_rows:
                    cur.execute("""
                        INSERT INTO history (sheet_name, role, user_id, snapshot)
                        VALUES (%s, %s, %s, %s)
                    """, (sn, 'tham_tra', tt_row['user_id'], snapshot_json))
                    count += 1

            conn.commit()
            print(f"📸 [SNAPSHOT] Đã tạo {count} history entry trước khi reset.")
    except Exception as e:
        print(f"❌ [SNAPSHOT] Lỗi khi tạo snapshot: {e}")


def reset_current_data():
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM evaluations")
        cur.execute("DELETE FROM review_comments")
        cur.execute("DELETE FROM suggestions")
        # KHÔNG xóa history — đây là lịch sử vĩnh viễn theo tháng
        conn.commit()
    print("🗑️ Đã xóa evaluations, comments, suggestions. History giữ nguyên.")

# ================== AUTO RESET KHI SANG THÁNG MỚI ==================
def ensure_app_config_table():
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS app_config (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)
        conn.commit()

def auto_reset_if_new_month(force=False):
    try:
        ensure_app_config_table()
        current_month = datetime.now().strftime('%Y-%m')

        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT value FROM app_config WHERE key = 'last_reset_month'")
            row = cur.fetchone()
            last_reset = row['value'] if row else None

        if last_reset is None:
            with get_db_connection() as conn:
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO app_config (key, value) VALUES ('last_reset_month', %s)
                    ON CONFLICT (key) DO NOTHING
                """, (current_month,))
            print("DB mới — đã ghi nhận tháng, không reset.")
            return False

        if force or last_reset != current_month:
            print(f"🔄 [AUTO RESET] Phát hiện tháng mới: {current_month}. Đang tiến hành reset...")

            snapshot_all_before_reset()
            archive_current_data()
            reset_current_data()

            with get_db_connection() as conn:
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO app_config (key, value)
                    VALUES ('last_reset_month', %s)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
                """, (current_month,))

            print(f"✅ [AUTO RESET] Hoàn tất reset tháng {current_month}.")
            return True

        return False

    except Exception as e:
        print(f"❌ [AUTO RESET] Lỗi hệ thống: {e}")
        return False


def _background_reset_check():
    """Chạy kiểm tra reset trong background thread — không block bất kỳ request nào"""
    try:
        auto_reset_if_new_month()
    except Exception as e:
        print(f"❌ [BACKGROUND RESET] Lỗi: {e}")

@app.before_request
def before_request_hook():
    current_time = time.time()
    # Kiểm tra mỗi 1 giờ, cập nhật timestamp TRƯỚC khi spawn thread
    # để tránh nhiều request cùng lúc tạo nhiều thread trùng nhau
    if not hasattr(app, '_last_reset_check') or (current_time - app._last_reset_check > 3600):
        app._last_reset_check = current_time
        t = threading.Thread(target=_background_reset_check, daemon=True)
        t.start()

# Chạy reset check một lần khi app khởi động (background, không block)
if not os.environ.get('TPM_DISABLE_AUTO_RESET'):
    _startup_thread = threading.Thread(target=_background_reset_check, daemon=True)
    _startup_thread.start()


# -------------------- ROUTES --------------------
@app.route('/')
def index():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM users WHERE username = %s AND password = %s", (username, password))
            user = cur.fetchone()
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

    current_month = datetime.now().strftime('%Y-%m')
    selected_month = request.args.get('month', current_month)
    is_current_month = (selected_month == current_month)

    with get_db_connection() as conn:
        cur = conn.cursor()

        cur.execute("""
            SELECT DISTINCT TO_CHAR(saved_at, 'YYYY-MM') as month
            FROM history
            WHERE user_id = %s
            ORDER BY month DESC
        """, (session['user_id'],))
        history_months = [r['month'] for r in cur.fetchall()]

        from datetime import timedelta
        available_months = []
        for i in range(6):
            dt = datetime.now() - timedelta(days=30 * i)
            available_months.append(dt.strftime('%Y-%m'))

        available_months = sorted(set(available_months + history_months), reverse=True)

        month_labels = {}
        for m in available_months:
            y, mo = m.split('-')
            month_labels[m] = f'Tháng {int(mo)}/{y}'

        cur.execute("""
            SELECT a.sheet_name, a.role,
                   COALESCE(s.locked_danh_gia, 0) as locked_danh_gia,
                   COALESCE(s.locked_tham_tra, 0) as locked_tham_tra
            FROM assignments a
            LEFT JOIN suggestions s ON a.sheet_name = s.sheet_name
            WHERE a.user_id = %s
        """, (session['user_id'],))
        assigns = cur.fetchall()

        if not is_current_month:
            cur.execute("""
                SELECT DISTINCT ON (sheet_name, role) id, sheet_name, role
                FROM history
                WHERE user_id = %s AND TO_CHAR(saved_at, 'YYYY-MM') = %s
                ORDER BY sheet_name, role, saved_at DESC
            """, (session['user_id'], selected_month))
            hist_rows = cur.fetchall()
            history_sheets = {(r['sheet_name'], r['role']) for r in hist_rows}
            history_ids = {(r['sheet_name'], r['role']): r['id'] for r in hist_rows}
        else:
            history_sheets = set()
            history_ids = {}

    eval_status = {}
    if is_current_month:
        tham_tra_sheets = [
        ass['sheet_name']
        for ass in assigns
        if ass['role'] == 'tham_tra'
    ]
    if tham_tra_sheets:
        # Default tất cả = False trước
        eval_status = {sheet: False for sheet in tham_tra_sheets}
        
        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT
                    e.sheet_name,
                    COUNT(*) AS total,
                    MAX(
                        CASE
                            WHEN s.reviewer_signature IS NOT NULL
                                 AND TRIM(s.reviewer_signature) != ''
                            THEN 1 ELSE 0
                        END
                    ) AS has_signature
                FROM evaluations e
                LEFT JOIN suggestions s ON e.sheet_name = s.sheet_name
                WHERE e.sheet_name = ANY(%s)
                  AND e.col_letter = 'G'
                  AND e.value IS NOT NULL
                  AND e.value != ''
                GROUP BY e.sheet_name
            """, (tham_tra_sheets,))
            rows = cur.fetchall()

        # Chỉ cập nhật những sheet có data thực tế
        eval_status.update({
            r['sheet_name']: (r['total'] > 0 and r['has_signature'] == 1)
            for r in rows
        })

    return render_template('dashboard.html',
        assignments=assigns,
        eval_status=eval_status,
        current_month=current_month,
        selected_month=selected_month,
        is_current_month=is_current_month,
        available_months=available_months,
        history_sheets=history_sheets,
        history_ids=history_ids,
        month_labels=month_labels
    )

@app.route('/form/<sheet_name>')
def evaluation_form(sheet_name):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT role FROM assignments WHERE user_id = %s AND sheet_name = %s", (session['user_id'], sheet_name))
        assign = cur.fetchone()
    if not assign:
        return "Bạn không có quyền truy cập biểu mẫu này", 403
    headers, rows, extra = get_sheet_data(sheet_name)
    if not headers:
        return f"Không tìm thấy sheet {sheet_name} trong file forms.xlsx", 404
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT row_index, col_letter, value FROM evaluations WHERE sheet_name = %s", (sheet_name,))
        db_rows = cur.fetchall()
        evals = {(r['row_index'], r['col_letter']): r['value'] for r in db_rows if r['row_index'] >= 10}
        saved_header = {(r['row_index'], r['col_letter']): r['value'] for r in db_rows if r['row_index'] < 10}
        cur.execute("SELECT row_index, comment FROM review_comments WHERE sheet_name = %s", (sheet_name,))
        comms = {r['row_index']: r['comment'] for r in cur.fetchall()}
        cur.execute("SELECT * FROM suggestions WHERE sheet_name = %s", (sheet_name,))
        s = cur.fetchone()
    suggestion = (s['suggestion'] if s else '') or ''
    reviewer_comment = (s['reviewer_comment'] if s else '') or ''
    reviewer_signature = (s['reviewer_signature'] if s else '') or ''
    checker_signature = (s['checker_signature'] if s else '') or ''
    locked_danh_gia = int(s['locked_danh_gia']) if s and s['locked_danh_gia'] is not None else 0
    locked_tham_tra = int(s['locked_tham_tra']) if s and s['locked_tham_tra'] is not None else 0
    return render_template('evaluation_form.html', sheet_name=sheet_name, role=assign['role'], headers=headers, rows=rows, extra=extra, saved=evals, saved_comments=comms, saved_header=saved_header, suggestion=suggestion, reviewer_comment=reviewer_comment, reviewer_signature=reviewer_signature, checker_signature=checker_signature, locked_danh_gia=locked_danh_gia, locked_tham_tra=locked_tham_tra, enumerate=enumerate)

@app.route('/evaluation/<sheet_name>')
def evaluation_redirect(sheet_name):
    return redirect(url_for('evaluation_form', sheet_name=sheet_name))

@app.route('/save', methods=['POST'])
def save():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    sn = request.form['sheet_name']
    role = request.form.get('role')
    uid = session['user_id']
    def col_name(col):
        return {'H': 'Mô tả', 'I': 'Đơn vị thực hiện', 'J': 'Thời gian', 'K': 'Giải pháp'}.get(col, col)
    with get_db_connection() as conn:
        cur = conn.cursor()
        if role == 'danh_gia':
            cycle_val = request.form.get('header_6_E', '').strip()
            if not cycle_val:
                flash('Vui lòng nhập "Nhiệt độ môi trường - Kiểm tra" (ô đầu tiên).')
                return redirect(url_for('evaluation_form', sheet_name=sn))
            cur.execute("""INSERT INTO evaluations (user_id, sheet_name, row_index, col_letter, value) VALUES (%s, %s, %s, %s, %s) ON CONFLICT (user_id, sheet_name, row_index, col_letter) DO UPDATE SET value = EXCLUDED.value""", (uid, sn, 6, 'E', cycle_val))
        elif role == 'tham_tra':
            cycle_val = request.form.get('header_6_F', '').strip()
            if not cycle_val:
                flash('Vui lòng nhập "Nhiệt độ môi trường - Thẩm tra" (ô thứ hai).')
                return redirect(url_for('evaluation_form', sheet_name=sn))
            cur.execute("""INSERT INTO evaluations (user_id, sheet_name, row_index, col_letter, value) VALUES (%s, %s, %s, %s, %s) ON CONFLICT (user_id, sheet_name, row_index, col_letter) DO UPDATE SET value = EXCLUDED.value""", (uid, sn, 6, 'F', cycle_val))
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
        for row, cols in eval_items.items():
            if 'G' not in cols or not cols['G'].strip():
                flash(f'Dòng {row}: chưa chọn kết quả (cột G).')
                return redirect(url_for('evaluation_form', sheet_name=sn))
            if cols.get('G') == 'K':
                missing = [col_name(c) for c in ['H','I','J','K'] if c not in cols or not cols[c].strip()]
                if missing:
                    flash(f'Dòng {row} (kết quả K) còn thiếu: {", ".join(missing)}.')
                    return redirect(url_for('evaluation_form', sheet_name=sn))
        reviewer_sig = request.form.get('reviewer_signature', '').strip()
        if not reviewer_sig:
            flash('Vui lòng nhập nội dung tại ô "Người đánh giá" (ký xác nhận).')
            return redirect(url_for('evaluation_form', sheet_name=sn))
        with get_db_connection() as conn:
            cur = conn.cursor()
            for k, v in request.form.items():
                if k.startswith('eval_'):
                    parts = k.split('_')
                    if len(parts) == 3:
                        try:
                            row = int(parts[1])
                            col = parts[2]
                        except:
                            continue
                        cur.execute("""INSERT INTO evaluations (user_id, sheet_name, row_index, col_letter, value) VALUES (%s, %s, %s, %s, %s) ON CONFLICT (user_id, sheet_name, row_index, col_letter) DO UPDATE SET value = EXCLUDED.value""", (uid, sn, row, col, v))
            now = datetime.now().strftime('%Hh%M ngày %d/%m/%y')
            cur.execute("""INSERT INTO evaluations (user_id, sheet_name, row_index, col_letter, value) VALUES (%s, %s, %s, %s, %s) ON CONFLICT (user_id, sheet_name, row_index, col_letter) DO UPDATE SET value = EXCLUDED.value""", (uid, sn, 4, 'F', now))
            cur.execute("""INSERT INTO suggestions (sheet_name, suggestion, reviewer_signature, locked_danh_gia) VALUES (%s, %s, %s, 1) ON CONFLICT (sheet_name) DO UPDATE SET suggestion = EXCLUDED.suggestion, reviewer_signature = EXCLUDED.reviewer_signature, locked_danh_gia = 1""", (sn, request.form.get('suggestion', ''), reviewer_sig))
            cur.execute("SELECT row_index, col_letter, value FROM evaluations WHERE sheet_name = %s", (sn,))
            evals_snap = cur.fetchall()
            cur.execute("SELECT row_index, comment FROM review_comments WHERE sheet_name = %s", (sn,))
            comms_snap = cur.fetchall()
            cur.execute("SELECT * FROM suggestions WHERE sheet_name = %s", (sn,))
            sugg_snap = cur.fetchone()
            dg_snapshot = json.dumps({
                'evals': [{'row': r['row_index'], 'col': r['col_letter'], 'value': r['value']} for r in evals_snap],
                'comments': [{'row': r['row_index'], 'comment': r['comment']} for r in comms_snap],
                'suggestions': dict(sugg_snap) if sugg_snap else {}
            }, ensure_ascii=False, default=str)
            cur.execute("DELETE FROM history WHERE sheet_name = %s AND role = 'danh_gia' AND user_id = %s", (sn, uid))
            cur.execute("INSERT INTO history (sheet_name, role, user_id, snapshot) VALUES (%s, %s, %s, %s)",
                        (sn, 'danh_gia', uid, dg_snapshot))
        flash('Đã lưu đánh giá thành công.')
    elif role == 'tham_tra':
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
        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT DISTINCT row_index FROM evaluations WHERE sheet_name = %s AND col_letter = 'G'", (sn,))
            rows_to_check = cur.fetchall()
        for r in rows_to_check:
            row = r['row_index']
            if row not in comment_items or not comment_items[row].strip():
                flash(f'Dòng {row}: chưa nhập ý kiến thẩm tra.')
                return redirect(url_for('evaluation_form', sheet_name=sn))
        checker_sig = request.form.get('checker_signature', '').strip()
        if not checker_sig:
            flash('Vui lòng nhập nội dung tại ô "Người thẩm tra" (ký xác nhận).')
            return redirect(url_for('evaluation_form', sheet_name=sn))
        with get_db_connection() as conn:
            cur = conn.cursor()
            for k, v in request.form.items():
                if k.startswith('comment_'):
                    parts = k.split('_')
                    if len(parts) == 2:
                        try:
                            row = int(parts[1])
                        except:
                            continue
                        cur.execute("""INSERT INTO review_comments (reviewer_id, sheet_name, row_index, comment) VALUES (%s, %s, %s, %s) ON CONFLICT (reviewer_id, sheet_name, row_index) DO UPDATE SET comment = EXCLUDED.comment""", (uid, sn, row, v))
            now = datetime.now().strftime('%Hh%M ngày %d/%m/%y')
            cur.execute("""INSERT INTO evaluations (user_id, sheet_name, row_index, col_letter, value) VALUES (%s, %s, %s, %s, %s) ON CONFLICT (user_id, sheet_name, row_index, col_letter) DO UPDATE SET value = EXCLUDED.value""", (uid, sn, 5, 'F', now))
            cur.execute("""INSERT INTO suggestions (sheet_name, reviewer_comment, checker_signature, locked_tham_tra) VALUES (%s, %s, %s, 1) ON CONFLICT (sheet_name) DO UPDATE SET reviewer_comment = EXCLUDED.reviewer_comment, checker_signature = EXCLUDED.checker_signature, locked_tham_tra = 1""", (sn, request.form.get('reviewer_comment', ''), checker_sig))
            cur.execute("SELECT row_index, col_letter, value FROM evaluations WHERE sheet_name = %s", (sn,))
            evals_snapshot = cur.fetchall()
            cur.execute("SELECT row_index, comment FROM review_comments WHERE sheet_name = %s", (sn,))
            comms_snapshot = cur.fetchall()
            cur.execute("SELECT * FROM suggestions WHERE sheet_name = %s", (sn,))
            sugg = cur.fetchone()
            snapshot = json.dumps({'evals': [{'row': r['row_index'], 'col': r['col_letter'], 'value': r['value']} for r in evals_snapshot], 'comments': [{'row': r['row_index'], 'comment': r['comment']} for r in comms_snapshot], 'suggestions': dict(sugg) if sugg else {}}, ensure_ascii=False)
            cur.execute("INSERT INTO history (sheet_name, role, user_id, snapshot) VALUES (%s, %s, %s, %s)", (sn, role, uid, snapshot))
        flash('Đã lưu ý kiến thẩm tra thành công.')
    return redirect(url_for('evaluation_form', sheet_name=sn))

# -------------------- Các route còn lại --------------------
@app.route('/history')
def history():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    selected_month = request.args.get('month')
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT TO_CHAR(saved_at, 'YYYY-MM') as month FROM history ORDER BY month DESC")
        months = cur.fetchall()
        if selected_month:
            cur.execute("SELECT h.id, h.sheet_name, h.role, h.saved_at, u.fullname FROM history h JOIN users u ON h.user_id = u.id WHERE TO_CHAR(h.saved_at, 'YYYY-MM') = %s ORDER BY h.saved_at DESC", (selected_month,))
            rows = cur.fetchall()
        else:
            cur.execute("SELECT h.id, h.sheet_name, h.role, h.saved_at, u.fullname FROM history h JOIN users u ON h.user_id = u.id ORDER BY h.saved_at DESC")
            rows = cur.fetchall()
    return render_template('history.html', history=rows, months=months, selected_month=selected_month)

@app.route('/view_history/<int:history_id>')
def view_history(history_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT h.*, u.fullname as user_fullname FROM history h JOIN users u ON h.user_id = u.id WHERE h.id = %s", (history_id,))
        h = cur.fetchone()
        if not h:
            flash('Không tìm thấy bản ghi lịch sử.')
            return redirect(url_for('history'))
        snapshot = json.loads(h['snapshot'])
        headers, rows, extra = get_sheet_data(h['sheet_name'])
        return render_template('view_history.html', history=h, snapshot=snapshot, headers=headers, rows=rows, extra=extra, enumerate=enumerate)

@app.route('/export_all_forms')
def export_all_forms():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền truy cập chức năng này.')
        return redirect(url_for('dashboard'))

    # Kết nối 1: lấy danh sách sheet + toàn bộ dữ liệu DB trong 1 lần duy nhất
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT sheet_name FROM suggestions WHERE locked_tham_tra = 1")
        sheets = cur.fetchall()

        if not sheets:
            flash('Chưa có biểu mẫu nào được thẩm tra hoàn thành.')
            return redirect(url_for('dashboard'))

        sheet_names = [s['sheet_name'] for s in sheets]

        cur.execute("""
            SELECT sheet_name, row_index, col_letter, value
            FROM evaluations
            WHERE sheet_name = ANY(%s)
        """, (sheet_names,))
        all_evals = {}
        for r in cur.fetchall():
            sn = r['sheet_name']
            if sn not in all_evals:
                all_evals[sn] = {}
            all_evals[sn][(r['row_index'], r['col_letter'])] = r['value']

        cur.execute("""
            SELECT sheet_name, row_index, comment
            FROM review_comments
            WHERE sheet_name = ANY(%s)
        """, (sheet_names,))
        all_comments = {}
        for r in cur.fetchall():
            sn = r['sheet_name']
            if sn not in all_comments:
                all_comments[sn] = {}
            all_comments[sn][r['row_index']] = r['comment']

        cur.execute("""
            SELECT * FROM suggestions
            WHERE sheet_name = ANY(%s)
        """, (sheet_names,))
        all_suggestions = {r['sheet_name']: r for r in cur.fetchall()}

    # Kết nối 2 (build_reverse_mapping đọc Excel, không cần DB thêm)
    rev_map = build_reverse_mapping()
    thin = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'),  bottom=Side(style='thin')
    )

    wb = Workbook()
    wb.remove(wb.active)

    for sheet in sheets:
        sheet_name = sheet['sheet_name']
        headers, rows, extra = get_sheet_data(sheet_name)
        if not headers:
            continue

        evals    = all_evals.get(sheet_name, {})
        comments = all_comments.get(sheet_name, {})
        sugg     = all_suggestions.get(sheet_name)

        display_name = rev_map.get(sheet_name, sheet_name)
        ws = wb.create_sheet(title=display_name[:31])

        for row in headers:
            ws.append([row.get(col, '') for col in 'ABCDEF'])
        ws.append([])
        ws.append(["Hạng mục", "STT", "Nội dung đánh giá", "Tiêu chuẩn",
                   "Phương pháp", "Trạng thái TB", "Kết quả", "Mô tả",
                   "Đơn vị thực hiện", "Thời gian", "Giải pháp"])

        for idx, row in enumerate(rows, start=10):
            ws.append([
                row['A'], row['B'], row['C'], row['D'], row['E'], row['F'],
                evals.get((idx, 'G'), ''), evals.get((idx, 'H'), ''),
                evals.get((idx, 'I'), ''), evals.get((idx, 'J'), ''),
                evals.get((idx, 'K'), '')
            ])
            if comments.get(idx):
                ws.cell(row=ws.max_row, column=12, value=comments[idx])

        ws.append([])
        ws.append(["Kiến nghị và ký xác nhận"])
        ws.append(["Kiến nghị (nếu có):", extra[0].get('B', '') if extra else ''])
        ws.append(["Người đánh giá:", sugg['reviewer_signature'] if sugg else ''])
        ws.append(["Người thẩm tra:", sugg['checker_signature'] if sugg else ''])

        for row in ws.iter_rows():
            for cell in row:
                cell.border = thin
        for col in ws.columns:
            max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        download_name=f'All_Forms_{datetime.now().strftime("%Y%m")}.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/reset_cycle')
def reset_cycle():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền truy cập chức năng này.')
        return redirect(url_for('dashboard'))
    return '''
    <!DOCTYPE html><html><head><title>Xác nhận reset chu kỳ</title></head>
    <body>
        <h2>Bạn có chắc chắn muốn reset dữ liệu cho chu kỳ mới?</h2>
        <p>Dữ liệu hiện tại sẽ được sao lưu và xóa để bắt đầu chu kỳ đánh giá mới.</p>
        <form method="post" action="/confirm_reset">
            <button type="submit">Xác nhận reset</button>
            <a href="/dashboard">Hủy</a>
        </form>
    </body></html>
    '''

@app.route('/confirm_reset', methods=['POST'])
def confirm_reset():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền truy cập chức năng này.')
        return redirect(url_for('dashboard'))
    snapshot_all_before_reset()
    archive_current_data()
    reset_current_data()
    flash('Đã sao lưu và reset dữ liệu cho chu kỳ mới.')
    return redirect(url_for('dashboard'))

@app.route('/export_summary')
def export_summary():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền truy cập.')
        return redirect(url_for('dashboard'))

    try:
        with get_db_connection() as conn:
            cur = conn.cursor(cursor_factory=RealDictCursor)
            # Dùng IN ('K', 'Đ') để lấy cả 2 trạng thái
            cur.execute("""
                SELECT e.sheet_name, e.row_index, e.value as result,
                       e2.value as description, rc.comment as reviewer_comment
                FROM evaluations e
                LEFT JOIN evaluations e2 ON e.sheet_name = e2.sheet_name AND e.row_index = e2.row_index AND e2.col_letter = 'H'
                LEFT JOIN review_comments rc ON e.sheet_name = rc.sheet_name AND e.row_index = rc.row_index
                WHERE e.col_letter = 'G' AND e.value IN ('K', 'Đ')
                ORDER BY e.sheet_name, e.row_index
            """)
            rows = cur.fetchall()
            
            # Lấy suggestions
            cur.execute("""
                SELECT sheet_name, reviewer_signature, checker_signature FROM suggestions
                WHERE sheet_name IN (SELECT DISTINCT sheet_name FROM evaluations WHERE col_letter = 'G' AND value IN ('K', 'Đ'))
            """)
            suggestions = cur.fetchall()
            sug_dict = {s['sheet_name']: (s['reviewer_signature'], s['checker_signature']) for s in suggestions}

        if not rows:
            flash('Không có dữ liệu (K hoặc Đ) để xuất báo cáo.')
            return redirect(url_for('dashboard'))

        rev_map = build_reverse_mapping()
        wb = load_workbook('template.xlsx') # Đảm bảo file template nằm cùng thư mục
        ws = wb.active
        
        # Chèn dòng để giữ định dạng
        num_rows = len(rows)
        if num_rows > 1:
            ws.insert_rows(3, amount=num_rows - 1)

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Đổ dữ liệu
        for stt, row in enumerate(rows, start=1):
            curr_row = 2 + stt 
            ws.cell(row=curr_row, column=1, value=stt).border = thin_border
            ws.cell(row=curr_row, column=2, value=rev_map.get(row['sheet_name'], row['sheet_name'])).border = thin_border
            ws.cell(row=curr_row, column=3, value=row['description'] or '').border = thin_border
            ws.cell(row=curr_row, column=4, value=row['reviewer_comment'] or '').border = thin_border

        # Đổ Kiến nghị
        base_ki_nghi_row = 6 + (num_rows - 1)
        ki_nghi_row = base_ki_nghi_row + 1
        
        for sc, (rs, cc) in sug_dict.items():
            ws.cell(row=ki_nghi_row, column=1, value=sc)
            ws.cell(row=ki_nghi_row, column=2, value="Kiến nghị:")
            ws.cell(row=ki_nghi_row, column=3, value=rs or '')
            ki_nghi_row += 1
            ws.cell(row=ki_nghi_row, column=2, value="Ý kiến thẩm tra:")
            ws.cell(row=ki_nghi_row, column=3, value=cc or '')
            ki_nghi_row += 2

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(output, download_name='TonghopKKTB_TPM.xlsx', as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        print(f"❌ Lỗi xuất file: {str(e)}")
        flash(f"Lỗi hệ thống: {str(e)}")
        return redirect(url_for('dashboard'))

@app.route('/admin_dashboard')
def admin_dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    user_id = session['user_id']
    is_admin = session.get('role') == 'admin'

    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT
                sheet_name,
                COUNT(*) as total,
                COUNT(CASE WHEN value = 'K' THEN 1 END) as k_count
            FROM evaluations
            WHERE col_letter = 'G'
              AND value IS NOT NULL
              AND value != ''
            GROUP BY sheet_name
        """)
        rows = cur.fetchall()

    system_data = []
    total_k = 0
    total_evals = 0

    for row in rows:
        total = row['total'] or 0
        k_count = row['k_count'] or 0
        percentage = round((k_count / total * 100) if total > 0 else 0, 1)
        system_data.append({
            'name': row['sheet_name'],
            'total': total,
            'k_count': k_count,
            'percentage': percentage
        })
        total_k += k_count
        total_evals += total

    avg_pct = round((total_k / total_evals * 100) if total_evals > 0 else 0, 1)
    total_ok = total_evals - total_k

    system_data.sort(key=lambda x: x['percentage'], reverse=True)

    return render_template('admin_dashboard.html',
                         system_data=system_data,
                         total_systems=len(system_data),
                         total_k=total_k,
                         avg_pct=avg_pct,
                         total_ok=total_ok,
                         is_admin=is_admin)

@app.route('/sync_assignments', methods=['POST'])
def sync_assignments():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền truy cập chức năng này.')
        return redirect(url_for('dashboard'))
    if not os.path.exists(PHAN_GIAO_FILE):
        flash('Không tìm thấy file phan_giao.xlsx')
        return redirect(url_for('dashboard'))
    mapping = build_sheet_mapping()
    new_assignments = []
    wb = safe_load_workbook(PHAN_GIAO_FILE)
    if not wb:
        flash('Không thể đọc file phan_giao.xlsx')
        return redirect(url_for('dashboard'))
    ws = wb.active
    for row in range(8, ws.max_row + 1):
        ma_bieu_mau = str(ws[f'C{row}'].value or '').strip()
        name_eval = str(ws[f'E{row}'].value or '').strip()
        name_check = str(ws[f'F{row}'].value or '').strip()
        if not ma_bieu_mau or (not name_eval and not name_check):
            continue
        if ma_bieu_mau in ['BM.P4.15.18', 'BM.P4.15.19']:
            base_num = ma_bieu_mau.split('.')[-1]
            snames = [f'BM{base_num}_a', f'BM{base_num}_b', f'BM{base_num}_c']
        else:
            sname = mapping.get(ma_bieu_mau)
            if not sname:
                continue
            snames = [sname]
        for sn in snames:
            if name_eval:
                new_assignments.append((name_eval, sn, 'danh_gia'))
            if name_check:
                new_assignments.append((name_check, sn, 'tham_tra'))
    wb.close()
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT a.id, a.sheet_name, a.role, u.fullname FROM assignments a JOIN users u ON a.user_id = u.id")
        current = cur.fetchall()
        current_set = {(r['fullname'], r['sheet_name'], r['role']) for r in current}
        new_set = set(new_assignments)
        added, removed = [], []
        for (fullname, sn, role) in new_set - current_set:
            uid = create_or_get_user(conn, fullname, role)
            cur.execute("INSERT INTO assignments (user_id, sheet_name, role) VALUES (%s, %s, %s)", (uid, sn, role))
            added.append(f"{fullname} - {sn} ({'đánh giá' if role=='danh_gia' else 'thẩm tra'})")
        for r in current:
            if (r['fullname'], r['sheet_name'], r['role']) not in new_set:
                cur.execute("DELETE FROM assignments WHERE id = %s", (r['id'],))
                removed.append(f"{r['fullname']} - {r['sheet_name']} ({'đánh giá' if r['role']=='danh_gia' else 'thẩm tra'})")
    flash('Đã đồng bộ phân công từ file phan_giao.xlsx.')
    if added:
        flash(f'✅ Thêm mới {len(added)} bản ghi:')
        for item in added[:20]:
            flash(f'  + {item}')
        if len(added) > 20:
            flash(f'  ... và {len(added)-20} bản ghi khác')
    if removed:
        flash(f'❌ Xóa bỏ {len(removed)} bản ghi:')
        for item in removed[:20]:
            flash(f'  - {item}')
        if len(removed) > 20:
            flash(f'  ... và {len(removed)-20} bản ghi khác')
    if not added and not removed:
        flash('Không có thay đổi nào so với phân công hiện tại.')
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/force_reset')
def force_reset():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền thực hiện thao tác này')
        return redirect(url_for('dashboard'))

    current_month = datetime.now().strftime('%Y-%m')
    print(f"🔴 FORCE RESET kích hoạt cho tháng: {current_month}")

    try:
        snapshot_all_before_reset()
        archive_current_data()
        reset_current_data()

        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO app_config (key, value)
                VALUES ('last_reset_month', %s)
                ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
            """, (current_month,))

        flash(f'✅ ĐÃ RESET TOÀN BỘ DỮ LIỆU để bắt đầu tháng {current_month}', 'success')
    except Exception as e:
        flash(f'❌ Lỗi reset: {str(e)}', 'danger')

    return redirect(url_for('dashboard'))

import atexit


@atexit.register
def close_db_pool():

    try:

        if db_pool:
            db_pool.closeall()
            print("🔌 Database pool closed")

    except Exception as e:

        print(f"❌ Lỗi đóng database pool: {e}")


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')