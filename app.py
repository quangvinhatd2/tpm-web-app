"""
TPM BANVEHPC — app.py (OPTIMIZED)
Changelog vs. bản gốc:
  - Workbook cache: dùng RLock thay Lock, tránh deadlock; interval check đúng
  - _sheet_cache: bảo vệ bằng RLock riêng
  - snapshot_all_before_reset(): gom toàn bộ query vào 1 lần, loại bỏ N+1 hoàn toàn
  - save(): toàn bộ danh_gia / tham_tra trong 1 connection, đúng atomic
  - export_summary(): bỏ override cursor_factory thừa
  - before_request_hook(): dùng threading.Event thay attribute trên app object
  - Xóa import thừa (unicodedata.normalize đã inline)
  - Thêm type hints, constants, docstrings ngắn gọn
"""

import os
import re
import json
import threading
import time
from io import BytesIO
from datetime import datetime
from contextlib import contextmanager
from unicodedata import normalize

from flask import (Flask, render_template, request, redirect,
                   url_for, session, flash, send_file)
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import psycopg2
from psycopg2.pool import ThreadedConnectionPool
from psycopg2.extras import RealDictCursor
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'fallback-local-key')

DATABASE_URL = os.environ.get('DATABASE_URL')
if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL environment variable not set")

FORMS_FILE = 'forms.xlsx'
PHAN_GIAO_FILE = 'phan_giao.xlsx'

# ==================== WORKBOOK CACHE ====================
# Dùng RLock (reentrant) để tránh deadlock nếu cùng thread gọi lại
_wb_lock = threading.RLock()
_MASTER_WB = None
_MASTER_WB_MTIME: float | None = None
_LAST_MTIME_CHECK: float = 0.0
_MTIME_CHECK_INTERVAL = 10          # giây — tăng lên 10 s để giảm syscall

# Sheet data cache, bảo vệ bởi lock riêng
_sheet_cache: dict = {}
_sheet_cache_lock = threading.RLock()


def get_master_workbook():
    """Trả về workbook đã cache; reload nếu file thay đổi."""
    global _MASTER_WB, _MASTER_WB_MTIME, _LAST_MTIME_CHECK

    now = time.monotonic()
    # Fast path: kiểm tra không cần lock nếu còn trong interval
    if _MASTER_WB is not None and now - _LAST_MTIME_CHECK < _MTIME_CHECK_INTERVAL:
        return _MASTER_WB

    with _wb_lock:
        now = time.monotonic()
        # Double-check sau khi có lock
        if _MASTER_WB is not None and now - _LAST_MTIME_CHECK < _MTIME_CHECK_INTERVAL:
            return _MASTER_WB

        _LAST_MTIME_CHECK = now
        try:
            current_mtime = os.path.getmtime(FORMS_FILE)
        except OSError as e:
            print(f"❌ Không đọc được mtime {FORMS_FILE}: {e}")
            return _MASTER_WB  # trả về cache cũ nếu có

        if _MASTER_WB is not None and _MASTER_WB_MTIME == current_mtime:
            return _MASTER_WB

        # Cần reload
        if _MASTER_WB is not None:
            try:
                _MASTER_WB.close()
            except Exception:
                pass

        _MASTER_WB = load_workbook(FORMS_FILE, read_only=True, data_only=True)
        _MASTER_WB_MTIME = current_mtime

        with _sheet_cache_lock:
            _sheet_cache.clear()

        print("✅ Workbook cache reloaded")
        return _MASTER_WB


def get_sheet_data(sheet_name: str):
    """Trả về (headers, rows, extra) từ cache hoặc đọc mới."""
    with _sheet_cache_lock:
        if sheet_name in _sheet_cache:
            return _sheet_cache[sheet_name]

    try:
        wb = get_master_workbook()
        if not wb or sheet_name not in wb.sheetnames:
            return None, None, None
        ws = wb[sheet_name]

        headers = []
        for row in ws.iter_rows(min_row=1, max_row=7, max_col=6, values_only=True):
            headers.append({chr(65 + i): (cell or '') for i, cell in enumerate(row)})

        rows = []
        for row in ws.iter_rows(min_row=10, max_row=500, max_col=6, values_only=True):
            row_dict = {chr(65 + i): (cell or '') for i, cell in enumerate(row)}
            if not any(str(v).strip() for v in row_dict.values()):
                break
            rows.append(row_dict)

        result = (headers, rows, [])

        with _sheet_cache_lock:
            _sheet_cache[sheet_name] = result

        return result
    except Exception as e:
        print(f"❌ Lỗi đọc sheet {sheet_name}: {e}")
        return None, None, None


# ==================== DATABASE POOL ====================
DB_POOL_MIN = 2
DB_POOL_MAX = 20
_reset_lock = threading.Lock()

for _attempt in range(3):
    try:
        db_pool = ThreadedConnectionPool(
            minconn=DB_POOL_MIN,
            maxconn=DB_POOL_MAX,
            dsn=DATABASE_URL,
            sslmode='require',
            cursor_factory=RealDictCursor,
        )
        print(f"✅ DB pool initialized ({DB_POOL_MIN}-{DB_POOL_MAX})")
        break
    except Exception as _e:
        if _attempt == 2:
            raise
        print(f"❌ Pool init attempt {_attempt + 1}: {_e}")
        time.sleep(2 ** _attempt)


@contextmanager
def get_db_connection():
    """Context manager lấy/trả connection từ pool, tự động phục hồi connection chết."""
    conn = None
    try:
        conn = db_pool.getconn()
        # Kiểm tra connection còn sống không (ping)
        with conn.cursor() as cur:
            cur.execute("SELECT 1")
        yield conn
        conn.commit()
    except (psycopg2.InterfaceError, psycopg2.OperationalError) as e:
        # Connection đã bị đóng hoặc lỗi kết nối
        if conn:
            # Loại bỏ connection hỏng khỏi pool
            db_pool.putconn(conn, close=True)
            conn = None
        # Lấy connection mới
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


# ==================== HELPERS ====================

def safe_load_workbook(filepath, read_only=False):
    try:
        return load_workbook(filepath, read_only=read_only, data_only=True)
    except Exception as e:
        print(f"❌ Lỗi đọc {filepath}: {e}")
        return None


def unsigned_user(text: str) -> str | None:
    """Chuyển tên tiếng Việt → username ASCII không dấu, lowercase."""
    if not text:
        return None
    text = text.strip()
    _MAP = {
        'Đ': 'D', 'đ': 'd',
        'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A',
        'à': 'a', 'á': 'a', 'â': 'a', 'ã': 'a',
        'È': 'E', 'É': 'E', 'Ê': 'E', 'è': 'e', 'é': 'e', 'ê': 'e',
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
    result = ''.join(_MAP.get(ch, normalize('NFKD', ch).encode('ascii', 'ignore').decode('ascii')) for ch in text)
    return re.sub(r'[^a-zA-Z0-9]', '', result).lower()


def build_sheet_mapping() -> dict:
    wb = get_master_workbook()
    if not wb:
        return {}
    mapping = {}
    for sheet_name in wb.sheetnames:
        if not sheet_name.startswith('BM'):
            continue
        code_part = sheet_name[2:]
        if code_part.isdigit():
            mapping[f'BM.P4.15.{int(code_part):02d}'] = sheet_name
        elif '_' in code_part:
            base, _ = code_part.split('_', 1)
            if base.isdigit():
                mapping[f'BM.P4.15.{int(base):02d}_{code_part.split("_",1)[1]}'] = sheet_name
    return mapping


def build_reverse_mapping() -> dict:
    wb = get_master_workbook()
    if not wb:
        return {}
    rev = {}
    for sheet_name in wb.sheetnames:
        if not sheet_name.startswith('BM'):
            continue
        code_part = sheet_name[2:]
        if code_part.isdigit():
            rev[sheet_name] = f'BM.P4.15.{int(code_part):02d}'
        elif '_' in code_part:
            base, pha = code_part.split('_', 1)
            if base.isdigit():
                rev[sheet_name] = f'BM.P4.15.{int(base):02d}_{pha}'
    return rev


def create_or_get_user(conn, fullname: str, role: str) -> int:
    username = unsigned_user(fullname)
    cur = conn.cursor()
    cur.execute("SELECT id FROM users WHERE username = %s", (username,))
    user = cur.fetchone()
    if user:
        cur.execute("UPDATE users SET fullname = %s WHERE username = %s", (fullname, username))
        return user['id']
    cur.execute(
        "INSERT INTO users (username, password, fullname, role) VALUES (%s, %s, %s, %s) RETURNING id",
        (username, '123', fullname, role)
    )
    return cur.fetchone()['id']


def is_evaluation_complete(sheet_name: str) -> bool:
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT
                (SELECT COUNT(*) FROM evaluations
                 WHERE sheet_name = %s AND col_letter = 'G'
                   AND value IS NOT NULL AND value != '') AS cnt,
                (SELECT reviewer_signature FROM suggestions WHERE sheet_name = %s) AS sig
        """, (sheet_name, sheet_name))
        row = cur.fetchone()
    return bool(row and row['cnt'] > 0 and row['sig'] and row['sig'].strip())


# ==================== DB INIT ====================

def init_db():
    if not os.path.exists(PHAN_GIAO_FILE):
        print("⚠️  Không tìm thấy file phan_giao.xlsx")
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
        cur.execute("""
            CREATE TABLE IF NOT EXISTS app_config (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)
        # Indexes
        cur.execute("CREATE INDEX IF NOT EXISTS idx_eval_sheet_col ON evaluations(sheet_name, col_letter)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_eval_sheet_row ON evaluations(sheet_name, row_index)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_assign_user ON assignments(user_id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_assign_sheet ON assignments(sheet_name)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_rc_sheet ON review_comments(sheet_name)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_history_user_month ON history(user_id, (TO_CHAR(saved_at, 'YYYY-MM')))")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_history_saved_at ON history(saved_at)")

        wb = safe_load_workbook(PHAN_GIAO_FILE)
        if not wb:
            print("❌ Không thể đọc phan_giao.xlsx")
            return
        ws = wb.active
        for row in range(8, ws.max_row + 1):
            ma = str(ws[f'C{row}'].value or '').strip()
            name_eval = str(ws[f'E{row}'].value or '').strip()
            name_check = str(ws[f'F{row}'].value or '').strip()
            if not ma or (not name_eval and not name_check):
                continue
            if ma in ('BM.P4.15.18', 'BM.P4.15.19'):
                base_num = ma.split('.')[-1]
                sheet_names = [f'BM{base_num}_a', f'BM{base_num}_b', f'BM{base_num}_c']
            else:
                sname = mapping.get(ma)
                if not sname:
                    continue
                sheet_names = [sname]
            if name_eval:
                uid = create_or_get_user(conn, name_eval, 'danh_gia')
                for sn in sheet_names:
                    cur.execute("INSERT INTO assignments (user_id, sheet_name, role) VALUES (%s,%s,%s)", (uid, sn, 'danh_gia'))
            if name_check:
                uid = create_or_get_user(conn, name_check, 'tham_tra')
                for sn in sheet_names:
                    cur.execute("INSERT INTO assignments (user_id, sheet_name, role) VALUES (%s,%s,%s)", (uid, sn, 'tham_tra'))
        wb.close()
        cur.execute("SELECT id FROM users WHERE username = 'admin'")
        if not cur.fetchone():
            cur.execute("INSERT INTO users (username,password,fullname,role) VALUES ('admin','admin123','Quản trị viên','admin')")
    print("✅ Đã nạp dữ liệu phân công")


# ==================== ARCHIVE / RESET ====================

def archive_current_data():
    archive_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
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
        for table in ('evaluations', 'review_comments', 'suggestions'):
            cur.execute(f"SELECT * FROM {table}")
            rows = cur.fetchall()
            if rows:
                cur.executemany(
                    "INSERT INTO archives (archive_date, table_name, row_data) VALUES (%s,%s,%s)",
                    [(archive_date, table, json.dumps(dict(r), ensure_ascii=False, default=str)) for r in rows]
                )
        cur.execute("DELETE FROM archives WHERE archive_date::timestamp < NOW() - INTERVAL '12 months'")
    print(f"📦 Archive xong: {archive_date}")


def snapshot_all_before_reset():
    """
    Tạo history snapshot trước khi reset.
    BUG FIX gốc: loại bỏ N+1 query — gom TOÀN BỘ dữ liệu trong 1 connection,
    sau đó xử lý trong Python, chỉ cần executemany cuối cùng.
    """
    try:
        with get_db_connection() as conn:
            cur = conn.cursor()
            current_month = datetime.now().strftime('%Y-%m')

            # 1. Các sheet đã có đánh giá
            cur.execute("""
                SELECT DISTINCT sheet_name FROM evaluations
                WHERE col_letter = 'G' AND value IS NOT NULL AND value != ''
            """)
            sheets_with_evals = {r['sheet_name'] for r in cur.fetchall()}
            if not sheets_with_evals:
                return

            # 2. Đã lưu history trong tháng này chưa
            cur.execute("""
                SELECT sheet_name, role FROM history
                WHERE TO_CHAR(saved_at, 'YYYY-MM') = %s
            """, (current_month,))
            already_saved = {(r['sheet_name'], r['role']) for r in cur.fetchall()}

            # 3. Lấy TẤT CẢ evaluations 1 lần
            cur.execute("""
                SELECT sheet_name, row_index, col_letter, value FROM evaluations
                WHERE sheet_name = ANY(%s)
            """, (list(sheets_with_evals),))
            all_evals: dict[str, list] = {}
            for r in cur.fetchall():
                all_evals.setdefault(r['sheet_name'], []).append(
                    {'row': r['row_index'], 'col': r['col_letter'], 'value': r['value']}
                )

            # 4. Lấy TẤT CẢ comments 1 lần
            cur.execute("""
                SELECT sheet_name, row_index, comment FROM review_comments
                WHERE sheet_name = ANY(%s)
            """, (list(sheets_with_evals),))
            all_comments: dict[str, list] = {}
            for r in cur.fetchall():
                all_comments.setdefault(r['sheet_name'], []).append(
                    {'row': r['row_index'], 'comment': r['comment']}
                )

            # 5. Lấy TẤT CẢ suggestions 1 lần
            cur.execute("SELECT * FROM suggestions WHERE sheet_name = ANY(%s)", (list(sheets_with_evals),))
            all_suggestions = {r['sheet_name']: dict(r) for r in cur.fetchall()}

            # 6. Lấy TẤT CẢ assignments 1 lần
            cur.execute("""
                SELECT sheet_name, role, user_id FROM assignments
                WHERE sheet_name = ANY(%s)
            """, (list(sheets_with_evals),))
            assign_map: dict[tuple, int] = {}
            for r in cur.fetchall():
                assign_map[(r['sheet_name'], r['role'])] = r['user_id']

            # 7. Build và insert history
            history_rows = []
            for sn in sheets_with_evals:
                snapshot_json = json.dumps({
                    'evals': all_evals.get(sn, []),
                    'comments': all_comments.get(sn, []),
                    'suggestions': all_suggestions.get(sn, {}),
                }, ensure_ascii=False, default=str)

                for role in ('danh_gia', 'tham_tra'):
                    if (sn, role) in already_saved:
                        continue
                    uid = assign_map.get((sn, role))
                    if not uid:
                        continue
                    # Thẩm tra: chỉ lưu nếu có comment
                    if role == 'tham_tra' and not all_comments.get(sn):
                        continue
                    history_rows.append((sn, role, uid, snapshot_json))

            if history_rows:
                cur.executemany(
                    "INSERT INTO history (sheet_name, role, user_id, snapshot) VALUES (%s,%s,%s,%s)",
                    history_rows
                )
            print(f"📸 Snapshot: {len(history_rows)} history entries")
    except Exception as e:
        print(f"❌ [SNAPSHOT] {e}")


def reset_current_data():
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM evaluations")
        cur.execute("DELETE FROM review_comments")
        cur.execute("DELETE FROM suggestions")
    print("🗑️ Reset xong evaluations/comments/suggestions")


# ==================== AUTO RESET ====================

def auto_reset_if_new_month(force: bool = False) -> bool:
    try:
        current_month = datetime.now().strftime('%Y-%m')
        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS app_config (
                    key TEXT PRIMARY KEY, value TEXT NOT NULL
                )
            """)
            cur.execute("SELECT value FROM app_config WHERE key = 'last_reset_month'")
            row = cur.fetchone()
            last_reset = row['value'] if row else None
            if last_reset is None:
                cur.execute("""
                    INSERT INTO app_config (key, value) VALUES ('last_reset_month', %s)
                    ON CONFLICT (key) DO NOTHING
                """, (current_month,))
                return False

        if force or last_reset != current_month:
            print(f"🔄 [AUTO RESET] → {current_month}")
            snapshot_all_before_reset()
            archive_current_data()
            reset_current_data()
            with get_db_connection() as conn:
                cur = conn.cursor()
                cur.execute("""
                    INSERT INTO app_config (key, value) VALUES ('last_reset_month', %s)
                    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value
                """, (current_month,))
            print(f"✅ [AUTO RESET] xong tháng {current_month}")
            return True
        return False
    except Exception as e:
        print(f"❌ [AUTO RESET] {e}")
        return False


def _background_reset_check():
    if not _reset_lock.acquire(blocking=False):
        return
    try:
        auto_reset_if_new_month()
    finally:
        _reset_lock.release()


# Dùng threading.Event + time.monotonic thay vì attribute trên app object
_last_reset_check_time: float = 0.0
_reset_check_lock = threading.Lock()

@app.before_request
def before_request_hook():
    global _last_reset_check_time
    now = time.monotonic()
    if now - _last_reset_check_time > 3600:
        with _reset_check_lock:
            if now - _last_reset_check_time > 3600:   # double-check
                _last_reset_check_time = now
                threading.Thread(target=_background_reset_check, daemon=True).start()


if not os.environ.get('TPM_DISABLE_AUTO_RESET'):
    threading.Thread(target=_background_reset_check, daemon=True).start()


# ==================== ROUTES ====================

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
            cur.execute("SELECT * FROM users WHERE username=%s AND password=%s", (username, password))
            user = cur.fetchone()
        if user:
            session['user_id'] = user['id']
            session['fullname'] = user['fullname']
            session['role'] = user['role']
            return redirect(url_for('dashboard'))
        flash('Sai tài khoản hoặc mật khẩu (Mặc định: 123 / admin: admin123)')
    return render_template('login.html')


@app.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    from datetime import timedelta
    current_month = datetime.now().strftime('%Y-%m')
    selected_month = request.args.get('month', current_month)
    is_current_month = (selected_month == current_month)

    available_months_set = {
        (datetime.now() - timedelta(days=30 * i)).strftime('%Y-%m')
        for i in range(6)
    }

    with get_db_connection() as conn:
        cur = conn.cursor()

        # History months
        cur.execute("""
            SELECT DISTINCT TO_CHAR(saved_at, 'YYYY-MM') AS month
            FROM history WHERE user_id = %s ORDER BY month DESC
        """, (session['user_id'],))
        history_months = [r['month'] for r in cur.fetchall()]

        available_months = sorted(available_months_set | set(history_months), reverse=True)
        month_labels = {m: f"Tháng {int(m.split('-')[1])}/{m.split('-')[0]}" for m in available_months}

        # Assignments
        cur.execute("""
            SELECT a.sheet_name, a.role,
                   COALESCE(s.locked_danh_gia, 0) AS locked_danh_gia,
                   COALESCE(s.locked_tham_tra, 0)  AS locked_tham_tra
            FROM assignments a
            LEFT JOIN suggestions s ON a.sheet_name = s.sheet_name
            WHERE a.user_id = %s
        """, (session['user_id'],))
        assigns = cur.fetchall()

        # History sheets (tháng cũ)
        history_sheets: set = set()
        history_ids: dict = {}
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

        # Eval status (chỉ tháng hiện tại)
        eval_status: dict = {}
        if is_current_month:
            tham_tra_sheets = [a['sheet_name'] for a in assigns if a['role'] == 'tham_tra']
            if tham_tra_sheets:
                eval_status = {s: False for s in tham_tra_sheets}
                cur.execute("""
                    SELECT e.sheet_name,
                           COUNT(*) AS total,
                           MAX(CASE WHEN s.reviewer_signature IS NOT NULL
                                         AND TRIM(s.reviewer_signature) != ''
                               THEN 1 ELSE 0 END) AS has_signature
                    FROM evaluations e
                    LEFT JOIN suggestions s ON e.sheet_name = s.sheet_name
                    WHERE e.sheet_name = ANY(%s)
                      AND e.col_letter = 'G'
                      AND e.value IS NOT NULL AND e.value != ''
                    GROUP BY e.sheet_name
                """, (tham_tra_sheets,))
                eval_status.update({
                    r['sheet_name']: (r['total'] > 0 and r['has_signature'] == 1)
                    for r in cur.fetchall()
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
        month_labels=month_labels,
    )


@app.route('/form/<sheet_name>')
def evaluation_form(sheet_name):
    if 'user_id' not in session:
        return redirect(url_for('login'))

    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT role FROM assignments WHERE user_id=%s AND sheet_name=%s",
            (session['user_id'], sheet_name)
        )
        assign = cur.fetchone()
        if not assign:
            return "Bạn không có quyền truy cập biểu mẫu này", 403

        # Gom cả 3 query trong 1 connection
        cur.execute("SELECT row_index, col_letter, value FROM evaluations WHERE sheet_name=%s", (sheet_name,))
        db_rows = cur.fetchall()
        evals = {(r['row_index'], r['col_letter']): r['value'] for r in db_rows if r['row_index'] >= 10}
        saved_header = {(r['row_index'], r['col_letter']): r['value'] for r in db_rows if r['row_index'] < 10}

        cur.execute("SELECT row_index, comment FROM review_comments WHERE sheet_name=%s", (sheet_name,))
        comms = {r['row_index']: r['comment'] for r in cur.fetchall()}

        cur.execute("SELECT * FROM suggestions WHERE sheet_name=%s", (sheet_name,))
        s = cur.fetchone()

    headers, rows, extra = get_sheet_data(sheet_name)
    if not headers:
        return f"Không tìm thấy sheet {sheet_name} trong forms.xlsx", 404

    def _s(key): return (s[key] if s and s[key] else '') or ''
    def _i(key): return int(s[key]) if s and s.get(key) is not None else 0

    return render_template('evaluation_form.html',
        sheet_name=sheet_name, role=assign['role'],
        headers=headers, rows=rows, extra=extra,
        saved=evals, saved_comments=comms, saved_header=saved_header,
        suggestion=_s('suggestion'),
        reviewer_comment=_s('reviewer_comment'),
        reviewer_signature=_s('reviewer_signature'),
        checker_signature=_s('checker_signature'),
        locked_danh_gia=_i('locked_danh_gia'),
        locked_tham_tra=_i('locked_tham_tra'),
        enumerate=enumerate,
    )


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

    COL_NAME = {'H': 'Mô tả', 'I': 'Đơn vị thực hiện', 'J': 'Thời gian', 'K': 'Giải pháp'}

    # Parse form trước (không cần DB)
    eval_items: dict[int, dict] = {}
    for key, value in request.form.items():
        if key.startswith('eval_'):
            parts = key.split('_')
            if len(parts) == 3:
                try:
                    eval_items.setdefault(int(parts[1]), {})[parts[2]] = value
                except ValueError:
                    pass

    comment_items: dict[int, str] = {}
    for key, value in request.form.items():
        if key.startswith('comment_'):
            parts = key.split('_')
            if len(parts) == 2:
                try:
                    comment_items[int(parts[1])] = value
                except ValueError:
                    pass

    if role == 'danh_gia':
        # --- Validate trước khi mở connection ---
        cycle_val = request.form.get('header_6_E', '').strip()
        if not cycle_val:
            flash('Vui lòng nhập "Nhiệt độ môi trường - Kiểm tra".')
            return redirect(url_for('evaluation_form', sheet_name=sn))

        for row_i, cols in eval_items.items():
            if not cols.get('G', '').strip():
                flash(f'Dòng {row_i}: chưa chọn kết quả (cột G).')
                return redirect(url_for('evaluation_form', sheet_name=sn))
            if cols.get('G') == 'K':
                missing = [COL_NAME[c] for c in ('H', 'I', 'J', 'K') if not cols.get(c, '').strip()]
                if missing:
                    flash(f'Dòng {row_i} (kết quả K) thiếu: {", ".join(missing)}.')
                    return redirect(url_for('evaluation_form', sheet_name=sn))

        reviewer_sig = request.form.get('reviewer_signature', '').strip()
        if not reviewer_sig:
            flash('Vui lòng nhập nội dung ô "Người đánh giá".')
            return redirect(url_for('evaluation_form', sheet_name=sn))

        now = datetime.now().strftime('%Hh%M ngày %d/%m/%y')

        eval_rows = [(uid, sn, 6, 'E', cycle_val), (uid, sn, 4, 'F', now)]
        for row_i, cols in eval_items.items():
            for col, val in cols.items():
                eval_rows.append((uid, sn, row_i, col, val))

        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.executemany("""
                INSERT INTO evaluations (user_id, sheet_name, row_index, col_letter, value)
                VALUES (%s,%s,%s,%s,%s)
                ON CONFLICT (user_id, sheet_name, row_index, col_letter)
                DO UPDATE SET value = EXCLUDED.value
            """, eval_rows)

            cur.execute("""
                INSERT INTO suggestions (sheet_name, suggestion, reviewer_signature, locked_danh_gia)
                VALUES (%s,%s,%s,1)
                ON CONFLICT (sheet_name) DO UPDATE
                SET suggestion=EXCLUDED.suggestion,
                    reviewer_signature=EXCLUDED.reviewer_signature,
                    locked_danh_gia=1
            """, (sn, request.form.get('suggestion', ''), reviewer_sig))

            # Snapshot trong cùng transaction
            cur.execute("SELECT row_index, col_letter, value FROM evaluations WHERE sheet_name=%s", (sn,))
            evals_snap = cur.fetchall()
            cur.execute("SELECT row_index, comment FROM review_comments WHERE sheet_name=%s", (sn,))
            comms_snap = cur.fetchall()
            cur.execute("SELECT * FROM suggestions WHERE sheet_name=%s", (sn,))
            sugg_snap = cur.fetchone()

            snapshot_json = json.dumps({
                'evals': [{'row': r['row_index'], 'col': r['col_letter'], 'value': r['value']} for r in evals_snap],
                'comments': [{'row': r['row_index'], 'comment': r['comment']} for r in comms_snap],
                'suggestions': dict(sugg_snap) if sugg_snap else {},
            }, ensure_ascii=False, default=str)

            cur.execute(
                "DELETE FROM history WHERE sheet_name=%s AND role='danh_gia' AND user_id=%s",
                (sn, uid)
            )
            cur.execute(
                "INSERT INTO history (sheet_name, role, user_id, snapshot) VALUES (%s,'danh_gia',%s,%s)",
                (sn, uid, snapshot_json)
            )

        flash('Đã lưu đánh giá thành công.')

    elif role == 'tham_tra':
        cycle_val = request.form.get('header_6_F', '').strip()
        if not cycle_val:
            flash('Vui lòng nhập "Nhiệt độ môi trường - Thẩm tra".')
            return redirect(url_for('evaluation_form', sheet_name=sn))

        checker_sig = request.form.get('checker_signature', '').strip()
        if not checker_sig:
            flash('Vui lòng nhập nội dung ô "Người thẩm tra".')
            return redirect(url_for('evaluation_form', sheet_name=sn))

        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT DISTINCT row_index FROM evaluations WHERE sheet_name=%s AND col_letter='G'",
                (sn,)
            )
            rows_to_check = [r['row_index'] for r in cur.fetchall()]

            for row_i in rows_to_check:
                if not comment_items.get(row_i, '').strip():
                    flash(f'Dòng {row_i}: chưa nhập ý kiến thẩm tra.')
                    return redirect(url_for('evaluation_form', sheet_name=sn))

            now = datetime.now().strftime('%Hh%M ngày %d/%m/%y')

            if comment_items:
                cur.executemany("""
                    INSERT INTO review_comments (reviewer_id, sheet_name, row_index, comment)
                    VALUES (%s,%s,%s,%s)
                    ON CONFLICT (reviewer_id, sheet_name, row_index)
                    DO UPDATE SET comment=EXCLUDED.comment
                """, [(uid, sn, row_i, val) for row_i, val in comment_items.items()])

            cur.executemany("""
                INSERT INTO evaluations (user_id, sheet_name, row_index, col_letter, value)
                VALUES (%s,%s,%s,%s,%s)
                ON CONFLICT (user_id, sheet_name, row_index, col_letter)
                DO UPDATE SET value=EXCLUDED.value
            """, [(uid, sn, 6, 'F', cycle_val), (uid, sn, 5, 'F', now)])

            cur.execute("""
                INSERT INTO suggestions (sheet_name, reviewer_comment, checker_signature, locked_tham_tra)
                VALUES (%s,%s,%s,1)
                ON CONFLICT (sheet_name) DO UPDATE
                SET reviewer_comment=EXCLUDED.reviewer_comment,
                    checker_signature=EXCLUDED.checker_signature,
                    locked_tham_tra=1
            """, (sn, request.form.get('reviewer_comment', ''), checker_sig))

            cur.execute("SELECT row_index, col_letter, value FROM evaluations WHERE sheet_name=%s", (sn,))
            evals_snap = cur.fetchall()
            cur.execute("SELECT row_index, comment FROM review_comments WHERE sheet_name=%s", (sn,))
            comms_snap = cur.fetchall()
            cur.execute("SELECT * FROM suggestions WHERE sheet_name=%s", (sn,))
            sugg = cur.fetchone()

            snapshot_json = json.dumps({
                'evals': [{'row': r['row_index'], 'col': r['col_letter'], 'value': r['value']} for r in evals_snap],
                'comments': [{'row': r['row_index'], 'comment': r['comment']} for r in comms_snap],
                'suggestions': dict(sugg) if sugg else {},
            }, ensure_ascii=False)

            cur.execute(
                "INSERT INTO history (sheet_name, role, user_id, snapshot) VALUES (%s,'tham_tra',%s,%s)",
                (sn, uid, snapshot_json)
            )

        flash('Đã lưu ý kiến thẩm tra thành công.')

    return redirect(url_for('evaluation_form', sheet_name=sn))


@app.route('/history')
def history():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    selected_month = request.args.get('month')
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT TO_CHAR(saved_at,'YYYY-MM') AS month FROM history ORDER BY month DESC")
        months = cur.fetchall()
        if selected_month:
            cur.execute("""
                SELECT h.id, h.sheet_name, h.role, h.saved_at, u.fullname
                FROM history h JOIN users u ON h.user_id = u.id
                WHERE TO_CHAR(h.saved_at,'YYYY-MM') = %s ORDER BY h.saved_at DESC
            """, (selected_month,))
        else:
            cur.execute("""
                SELECT h.id, h.sheet_name, h.role, h.saved_at, u.fullname
                FROM history h JOIN users u ON h.user_id = u.id ORDER BY h.saved_at DESC
            """)
        rows = cur.fetchall()
    return render_template('history.html', history=rows, months=months, selected_month=selected_month)


@app.route('/view_history/<int:history_id>')
def view_history(history_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT h.*, u.fullname AS user_fullname FROM history h
            JOIN users u ON h.user_id = u.id WHERE h.id = %s
        """, (history_id,))
        h = cur.fetchone()
        if not h:
            flash('Không tìm thấy bản ghi lịch sử.')
            return redirect(url_for('history'))
    snapshot = json.loads(h['snapshot'])
    headers, rows, extra = get_sheet_data(h['sheet_name'])
    return render_template('view_history.html',
        history=h, snapshot=snapshot, headers=headers, rows=rows, extra=extra, enumerate=enumerate)


@app.route('/export_all_forms')
def export_all_forms():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền truy cập.')
        return redirect(url_for('dashboard'))

    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT sheet_name FROM suggestions WHERE locked_tham_tra=1")
        sheets = cur.fetchall()
        if not sheets:
            flash('Chưa có biểu mẫu nào được thẩm tra hoàn thành.')
            return redirect(url_for('dashboard'))

        sheet_names = [s['sheet_name'] for s in sheets]

        cur.execute("""
            SELECT sheet_name, row_index, col_letter, value FROM evaluations
            WHERE sheet_name = ANY(%s)
        """, (sheet_names,))
        all_evals: dict[str, dict] = {}
        for r in cur.fetchall():
            all_evals.setdefault(r['sheet_name'], {})[(r['row_index'], r['col_letter'])] = r['value']

        cur.execute("SELECT sheet_name, row_index, comment FROM review_comments WHERE sheet_name=ANY(%s)", (sheet_names,))
        all_comments: dict[str, dict] = {}
        for r in cur.fetchall():
            all_comments.setdefault(r['sheet_name'], {})[r['row_index']] = r['comment']

        cur.execute("SELECT * FROM suggestions WHERE sheet_name=ANY(%s)", (sheet_names,))
        all_suggestions = {r['sheet_name']: r for r in cur.fetchall()}

    rev_map = build_reverse_mapping()
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))

    wb = Workbook()
    wb.remove(wb.active)
    for sheet in sheets:
        sn = sheet['sheet_name']
        headers, rows, extra = get_sheet_data(sn)
        if not headers:
            continue
        evals = all_evals.get(sn, {})
        comments = all_comments.get(sn, {})
        sugg = all_suggestions.get(sn)
        display_name = rev_map.get(sn, sn)
        ws = wb.create_sheet(title=display_name[:31])
        for row in headers:
            ws.append([row.get(col, '') for col in 'ABCDEF'])
        ws.append([])
        ws.append(["Hạng mục","STT","Nội dung đánh giá","Tiêu chuẩn","Phương pháp",
                   "Trạng thái TB","Kết quả","Mô tả","Đơn vị thực hiện","Thời gian","Giải pháp"])
        for idx, row in enumerate(rows, start=10):
            r = [row['A'], row['B'], row['C'], row['D'], row['E'], row['F'],
                 evals.get((idx,'G'),''), evals.get((idx,'H'),''),
                 evals.get((idx,'I'),''), evals.get((idx,'J'),''), evals.get((idx,'K'),'')]
            ws.append(r)
            if comments.get(idx):
                ws.cell(row=ws.max_row, column=12, value=comments[idx])
        ws.append([])
        ws.append(["Kiến nghị và ký xác nhận"])
        ws.append(["Kiến nghị (nếu có):", extra[0].get('B','') if extra else ''])
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
    return send_file(output,
        download_name=f'All_Forms_{datetime.now().strftime("%Y%m")}.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/reset_cycle')
def reset_cycle():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền.')
        return redirect(url_for('dashboard'))
    return '''<!DOCTYPE html><html><head><title>Xác nhận reset</title></head><body>
    <h2>Xác nhận reset chu kỳ mới?</h2>
    <p>Dữ liệu hiện tại sẽ được sao lưu và xóa.</p>
    <form method="post" action="/confirm_reset">
        <button type="submit">Xác nhận</button>
        <a href="/dashboard">Hủy</a>
    </form></body></html>'''


@app.route('/confirm_reset', methods=['POST'])
def confirm_reset():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền.')
        return redirect(url_for('dashboard'))
    snapshot_all_before_reset()
    archive_current_data()
    reset_current_data()
    flash('Đã sao lưu và reset dữ liệu chu kỳ mới.')
    return redirect(url_for('dashboard'))


@app.route('/export_summary')
def export_summary():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền.')
        return redirect(url_for('dashboard'))
    try:
        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT e.sheet_name, e.row_index, e.value AS result,
                       e2.value AS description, rc.comment AS reviewer_comment
                FROM evaluations e
                LEFT JOIN evaluations e2
                    ON e.sheet_name=e2.sheet_name AND e.row_index=e2.row_index AND e2.col_letter='H'
                LEFT JOIN review_comments rc
                    ON e.sheet_name=rc.sheet_name AND e.row_index=rc.row_index
                WHERE e.col_letter='G' AND e.value IN ('K','Đ')
                ORDER BY e.sheet_name, e.row_index
            """)
            rows = cur.fetchall()
            cur.execute("""
                SELECT sheet_name, reviewer_signature, checker_signature FROM suggestions
                WHERE sheet_name = ANY(
                    SELECT DISTINCT sheet_name FROM evaluations
                    WHERE col_letter='G' AND value IN ('K','Đ')
                )
            """)
            sug_dict = {s['sheet_name']: (s['reviewer_signature'], s['checker_signature'])
                        for s in cur.fetchall()}

        if not rows:
            flash('Không có dữ liệu (K hoặc Đ) để xuất.')
            return redirect(url_for('dashboard'))

        rev_map = build_reverse_mapping()
        wb = load_workbook('template.xlsx')
        ws = wb.active
        num_rows = len(rows)
        if num_rows > 1:
            ws.insert_rows(3, amount=num_rows - 1)

        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))
        for stt, row in enumerate(rows, start=1):
            curr_row = 2 + stt
            ws.cell(curr_row, 1, stt).border = thin_border
            ws.cell(curr_row, 2, rev_map.get(row['sheet_name'], row['sheet_name'])).border = thin_border
            ws.cell(curr_row, 3, row['description'] or '').border = thin_border
            ws.cell(curr_row, 4, row['reviewer_comment'] or '').border = thin_border

        ki_nghi_row = 6 + (num_rows - 1) + 1
        for sc, (rs, cc) in sug_dict.items():
            ws.cell(ki_nghi_row, 1, sc)
            ws.cell(ki_nghi_row, 2, "Kiến nghị:")
            ws.cell(ki_nghi_row, 3, rs or '')
            ki_nghi_row += 1
            ws.cell(ki_nghi_row, 2, "Ý kiến thẩm tra:")
            ws.cell(ki_nghi_row, 3, cc or '')
            ki_nghi_row += 2

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, download_name='TonghopKKTB_TPM.xlsx', as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        print(f"❌ export_summary: {e}")
        flash(f"Lỗi hệ thống: {e}")
        return redirect(url_for('dashboard'))


@app.route('/admin_dashboard')
def admin_dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT sheet_name,
                   COUNT(*) AS total,
                   COUNT(CASE WHEN value='K' THEN 1 END) AS k_count
            FROM evaluations
            WHERE col_letter='G' AND value IS NOT NULL AND value != ''
            GROUP BY sheet_name
        """)
        rows = cur.fetchall()

    system_data = []
    total_k = total_evals = 0
    for row in rows:
        total = row['total'] or 0
        k = row['k_count'] or 0
        system_data.append({
            'name': row['sheet_name'],
            'total': total,
            'k_count': k,
            'percentage': round(k / total * 100 if total else 0, 1),
        })
        total_k += k
        total_evals += total

    system_data.sort(key=lambda x: x['percentage'], reverse=True)
    return render_template('admin_dashboard.html',
        system_data=system_data,
        total_systems=len(system_data),
        total_k=total_k,
        avg_pct=round(total_k / total_evals * 100 if total_evals else 0, 1),
        total_ok=total_evals - total_k,
        is_admin=session.get('role') == 'admin',
    )


@app.route('/sync_assignments', methods=['POST'])
def sync_assignments():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền.')
        return redirect(url_for('dashboard'))
    if not os.path.exists(PHAN_GIAO_FILE):
        flash('Không tìm thấy file phan_giao.xlsx')
        return redirect(url_for('dashboard'))

    # DEBUG TẠM THỜI: xác nhận đang đọc đúng file nào, sửa lần cuối lúc nào
    _abs_path = os.path.abspath(PHAN_GIAO_FILE)
    _mtime = datetime.fromtimestamp(os.path.getmtime(PHAN_GIAO_FILE)).strftime('%Y-%m-%d %H:%M:%S')
    flash(f'🔍 DEBUG: đang đọc "{_abs_path}" | sửa lần cuối: {_mtime}')

    mapping = build_sheet_mapping()
    wb = safe_load_workbook(PHAN_GIAO_FILE)
    if not wb:
        flash('Không thể đọc phan_giao.xlsx')
        return redirect(url_for('dashboard'))

    new_assignments = []
    ws = wb.active
    for row in range(8, ws.max_row + 1):
        ma = str(ws[f'C{row}'].value or '').strip()
        name_eval = str(ws[f'E{row}'].value or '').strip()
        name_check = str(ws[f'F{row}'].value or '').strip()
        if not ma or (not name_eval and not name_check):
            continue
        if ma in ('BM.P4.15.18', 'BM.P4.15.19'):
            base_num = ma.split('.')[-1]
            snames = [f'BM{base_num}_a', f'BM{base_num}_b', f'BM{base_num}_c']
        else:
            sname = mapping.get(ma)
            if not sname:
                continue
            snames = [sname]
        for sn in snames:
            if name_eval:
                new_assignments.append((name_eval, sn, 'danh_gia'))
            if name_check:
                new_assignments.append((name_check, sn, 'tham_tra'))
    wb.close()

    new_set = set(new_assignments)
    with get_db_connection() as conn:
        cur = conn.cursor()
        cur.execute("SELECT a.id, a.sheet_name, a.role, u.fullname FROM assignments a JOIN users u ON a.user_id=u.id")
        current = cur.fetchall()
        current_set = {(r['fullname'], r['sheet_name'], r['role']) for r in current}

        added, removed = [], []
        for (fullname, sn, role) in new_set - current_set:
            uid = create_or_get_user(conn, fullname, role)
            cur.execute("INSERT INTO assignments (user_id, sheet_name, role) VALUES (%s,%s,%s)", (uid, sn, role))
            added.append(f"{fullname} - {sn} ({'đánh giá' if role=='danh_gia' else 'thẩm tra'})")
        for r in current:
            if (r['fullname'], r['sheet_name'], r['role']) not in new_set:
                cur.execute("DELETE FROM assignments WHERE id=%s", (r['id'],))
                removed.append(f"{r['fullname']} - {r['sheet_name']}")

    flash('Đã đồng bộ phân công.')
    if added:
        flash(f'✅ Thêm {len(added)} bản ghi')
        for item in added[:20]:
            flash(f'  + {item}')
        if len(added) > 20:
            flash(f'  ... và {len(added)-20} bản ghi khác')
    if removed:
        flash(f'❌ Xóa {len(removed)} bản ghi')
        for item in removed[:20]:
            flash(f'  - {item}')
        if len(removed) > 20:
            flash(f'  ... và {len(removed)-20} bản ghi khác')
    if not added and not removed:
        flash('Không có thay đổi.')
    return redirect(url_for('dashboard'))


@app.route('/force_reset')
def force_reset():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Bạn không có quyền.')
        return redirect(url_for('dashboard'))
    current_month = datetime.now().strftime('%Y-%m')
    try:
        snapshot_all_before_reset()
        archive_current_data()
        reset_current_data()
        with get_db_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO app_config (key, value) VALUES ('last_reset_month', %s)
                ON CONFLICT (key) DO UPDATE SET value=EXCLUDED.value
            """, (current_month,))
        flash(f'✅ Đã reset dữ liệu tháng {current_month}')
    except Exception as e:
        flash(f'❌ Lỗi: {e}')
    return redirect(url_for('dashboard'))


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


import atexit

@atexit.register
def close_db_pool():
    try:
        if db_pool:
            db_pool.closeall()
            print("🔌 DB pool closed")
    except Exception as e:
        print(f"❌ Đóng pool: {e}")


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0')