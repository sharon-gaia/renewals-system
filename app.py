import sys
import io
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', line_buffering=True)
except Exception:
    pass

from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_file
from flask_cors import CORS
import sqlite3
import os
import io
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook
from openpyxl import Workbook as NewWorkbook
import datetime
import imaplib
import email as email_lib
from email.header import decode_header
import threading
import time
import re
import pdfplumber
from bidi.algorithm import get_display
from dotenv import load_dotenv

load_dotenv()

# ── Email polling config ─────────────────────────────────────
EMAIL_CONFIG = {
    'imap_server': 'imap.gmail.com',
    'imap_port': 993,
    'username': os.environ['EMAIL_USERNAME'],
    'password': os.environ['EMAIL_PASSWORD'],
    'sender_filter': 'onboarding@resend.dev',
    'subject_filter': '',
    'check_interval': 300,
    'enabled': True,
}

app = Flask(__name__)
app.secret_key = os.environ['FLASK_SECRET_KEY']

from health_check import health_bp
app.register_blueprint(health_bp)
CORS(app, resources={r"/api/*": {"origins": [
    "https://www.winner-ins.co.il",
    "https://winner-ins.co.il",
    "https://www.gaia-ins.co.il",
    "https://gaia-ins.co.il"
]}})
DB_PATH = os.environ.get('DB_PATH', os.path.join(os.path.dirname(__file__), 'renewals.db')).strip()
_db_dir = os.path.dirname(DB_PATH)
if _db_dir:
    os.makedirs(_db_dir, exist_ok=True)
print(f'[startup] DB_PATH={DB_PATH}')

@app.template_filter('fdate')
def format_date(value):
    if not value:
        return '—'
    s = str(value).strip()
    # YYYY-MM-DD HH:MM or YYYY-MM-DD HH:MM:SS
    if len(s) >= 10 and s[4] == '-':
        parts = s.split(' ', 1)
        d = parts[0].split('-')
        if len(d) == 3:
            result = f"{d[2]}/{d[1]}/{d[0]}"
            if len(parts) > 1:
                result += ' ' + parts[1][:5]
            return result
    return s

STATUSES = ['', 'טופס התקבל', 'חודש', 'לא רוצים לחדש', 'לקוח ענה/ V כחול']
BRANDS = ['גאיה', 'ווינר', 'אופיר']

def normalize_id_number(s):
    """Israeli ID numbers are 9 digits — left-pad short numeric IDs with zeros
    (e.g. 33775065 → 033775065). Leaves non-numeric or 9+ digit values untouched."""
    s = str(s or '').strip()
    if s.isdigit() and len(s) < 9:
        return s.zfill(9)
    return s

def parse_dmy(s):
    """Parse a DD/MM/YYYY date string (as extracted from Harel PDFs) to a date. None on failure."""
    s = str(s or '').strip()
    m = re.match(r'(\d{2})/(\d{2})/(\d{4})', s)
    if not m:
        return None
    try:
        return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None

def brand_from_agency(agency):
    """Derive the brand (גאיה/ווינר/אופיר) from the agency name on the policy."""
    a = str(agency or '')
    if 'גאיה' in a:
        return 'גאיה'
    if 'ווינר' in a or 'וינר' in a:
        return 'ווינר'
    if 'אופיר' in a:
        return 'אופיר'
    return ''

def compute_active_status(period_end):
    """Active if today is on/before the policy end date; inactive once it has passed."""
    end = parse_dmy(period_end)
    if not end:
        return 'פעיל'  # unknown end date — assume active until told otherwise
    return 'פעיל' if datetime.date.today() <= end else 'לא פעיל'

def recompute_insured_statuses(conn):
    """Weekly job: refresh פעיל/לא פעיל by date. Never touches admin-overridden rows
    or ones already marked בוטל."""
    changed = 0
    for r in conn.execute(
        "SELECT id, period_end, status FROM insureds WHERE status_override=0 AND status != 'בוטל'"
    ).fetchall():
        new_status = compute_active_status(r['period_end'])
        if new_status != r['status']:
            conn.execute("UPDATE insureds SET status=?, updated_at=? WHERE id=?",
                         (new_status, datetime.datetime.now().isoformat(), r['id']))
            changed += 1
    conn.commit()
    return changed

def _event_sort_key(r):
    """Order a policy_records row on the timeline: prefer the document date, then period start."""
    d = str(r['doc_date'] or '')
    ps = parse_dmy(r['period_start'])
    return (d, ps.isoformat() if ps else '')

def rebuild_insureds(conn):
    """Build/refresh the insureds master from policy_records — one row per ID number,
    using each person's LATEST policy event (by document date). If that latest event is
    a cancellation (ביטול) the insured is marked 'בוטל'; otherwise status is by policy
    period. A stand-alone cancellation with no prior policy still creates the insured
    from the cancellation's own details. Preserves existing activity and admin overrides."""
    # Group policy_records by normalized ID, keep the latest event on the timeline
    best = {}
    for r in conn.execute(
        "SELECT * FROM policy_records WHERE insured_id IS NOT NULL AND insured_id != ''"
    ).fetchall():
        idn = normalize_id_number(r['insured_id'])
        if not idn:
            continue
        k = _event_sort_key(r)
        if idn not in best or k >= best[idn][0]:
            best[idn] = (k, r)

    now = datetime.datetime.now().isoformat()
    upserted = 0
    for idn, (_, r) in best.items():
        agency = r['agent_name'] or ''
        brand = brand_from_agency(agency)
        wa_source = 'ווינר' if brand in ('ווינר', 'אופיר') else None
        existing = conn.execute("SELECT id, status_override FROM insureds WHERE id_number=?", (idn,)).fetchone()
        # Cancellation wins when it is the latest event; otherwise status by period
        if 'ביטול' in str(r['doc_type_label'] or ''):
            status = 'בוטל'
        else:
            status = compute_active_status(r['period_end'])
        if existing:
            # Refresh policy/contact facts but never clobber activity or an admin override
            keep_status = existing['status_override'] == 1
            conn.execute(
                """UPDATE insureds SET name=?, agency=?, brand=?, phone=?, email=?, address=?,
                   policy_number=?, period_start=?, period_end=?, whatsapp_source=COALESCE(whatsapp_source, ?),
                   status=CASE WHEN status_override=1 THEN status ELSE ? END, updated_at=?
                   WHERE id=?""",
                (r['insured_name'], agency, brand, r['phone_mobile'] or r['phone_home'] or '',
                 r['email'], r['address'], r['policy_number'], r['period_start'], r['period_end'],
                 wa_source, status, now, existing['id'])
            )
        else:
            conn.execute(
                """INSERT INTO insureds
                   (id_number, name, agency, brand, phone, email, address, policy_number,
                    period_start, period_end, status, whatsapp_source, created_at, updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (idn, r['insured_name'], agency, brand, r['phone_mobile'] or r['phone_home'] or '',
                 r['email'], r['address'], r['policy_number'], r['period_start'], r['period_end'],
                 status, wa_source, now, now)
            )
        upserted += 1
    conn.commit()
    return upserted

def promote_customers_to_insureds(conn, month_id):
    """Move a renewal month's customers into the insureds master (req 4), preserving
    all activity (calls, notes, VIP, rep credit). Renewed → פעיל, otherwise → לא פעיל
    (req 5). Non-destructive: the original customers rows are left intact as history."""
    now = datetime.datetime.now().isoformat()
    promoted = 0
    for cst in conn.execute("SELECT * FROM customers WHERE month_id=?", (month_id,)).fetchall():
        idn = normalize_id_number(cst['id_number'])
        if not idn:
            continue
        status = 'פעיל' if cst['status'] == 'חודש' else 'לא פעיל'
        existing = conn.execute("SELECT * FROM insureds WHERE id_number=?", (idn,)).fetchone()
        if existing:
            # Fill blanks and set renewal-based status; never wipe existing activity.
            insured_has_calls = bool(existing['call_status_1'] or existing['call_status_2'] or existing['call_status_3'])
            conn.execute(
                """UPDATE insureds SET
                   name=COALESCE(NULLIF(name,''), ?),
                   phone=COALESCE(NULLIF(phone,''), ?),
                   brand=COALESCE(NULLIF(brand,''), ?),
                   whatsapp_source=COALESCE(whatsapp_source, ?),
                   agent_notes=COALESCE(NULLIF(agent_notes,''), ?),
                   is_vip=MAX(COALESCE(is_vip,0), ?),
                   handled_by=COALESCE(NULLIF(handled_by,''), ?),
                   policy_number=COALESCE(NULLIF(policy_number,''), ?),
                   status=?, updated_at=?
                   WHERE id=?""",
                (cst['name'], cst['phone'], cst['brand'], cst['whatsapp_source'],
                 cst['agent_notes'], cst['is_vip'] or 0, cst['handled_by'], cst['policy_number'],
                 status, now, existing['id'])
            )
            if not insured_has_calls:
                conn.execute(
                    """UPDATE insureds SET call_date_1=?, call_status_1=?, call_by_1=?,
                       call_date_2=?, call_status_2=?, call_by_2=?,
                       call_date_3=?, call_status_3=?, call_by_3=? WHERE id=?""",
                    (cst['call_date_1'], cst['call_status_1'], cst['call_by_1'],
                     cst['call_date_2'], cst['call_status_2'], cst['call_by_2'],
                     cst['call_date_3'], cst['call_status_3'], cst['call_by_3'], existing['id'])
                )
        else:
            wa_source = cst['whatsapp_source'] or ('ווינר' if cst['brand'] in ('ווינר', 'אופיר') else None)
            conn.execute(
                """INSERT INTO insureds
                   (id_number, name, phone, brand, whatsapp_source, agent_notes, is_vip, handled_by,
                    policy_number, status,
                    call_date_1, call_status_1, call_by_1, call_date_2, call_status_2, call_by_2,
                    call_date_3, call_status_3, call_by_3, created_at, updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (idn, cst['name'], cst['phone'], cst['brand'], wa_source, cst['agent_notes'],
                 cst['is_vip'] or 0, cst['handled_by'], cst['policy_number'], status,
                 cst['call_date_1'], cst['call_status_1'], cst['call_by_1'],
                 cst['call_date_2'], cst['call_status_2'], cst['call_by_2'],
                 cst['call_date_3'], cst['call_status_3'], cst['call_by_3'], now, now)
            )
        promoted += 1
    conn.commit()
    return promoted

# ── DB helpers ──────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            display_name TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'agent'
        );
        CREATE TABLE IF NOT EXISTS months (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month_id INTEGER NOT NULL,
            policy_number TEXT,
            name TEXT NOT NULL,
            id_number TEXT,
            phone TEXT,
            brand TEXT,
            status TEXT DEFAULT '',
            premium_last_year TEXT,
            whatsapp_sent_date TEXT,
            sharon_notes TEXT,
            requests_to_sharon TEXT,
            contact_date TEXT,
            agent_notes TEXT,
            interested_in_products TEXT,
            FOREIGN KEY (month_id) REFERENCES months(id)
        );
    ''')
    # Add form columns if missing (migration)
    existing = [r[1] for r in conn.execute("PRAGMA table_info(customers)").fetchall()]
    for col, typ in [('form_email','TEXT'), ('form_installments','TEXT'),
                     ('form_payment_method','TEXT'), ('form_received_at','TEXT'),
                     ('form_coverage','TEXT'), ('form_comments','TEXT'),
                     ('is_vip','INTEGER DEFAULT 0'), ('whatsapp_source','TEXT'),
                     ('call_date_1','TEXT'), ('call_status_1','TEXT'), ('call_by_1','TEXT'),
                     ('call_date_2','TEXT'), ('call_status_2','TEXT'), ('call_by_2','TEXT'),
                     ('call_date_3','TEXT'), ('call_status_3','TEXT'), ('call_by_3','TEXT')]:
        if col not in existing:
            conn.execute(f"ALTER TABLE customers ADD COLUMN {col} {typ}")
    # One-time backfill: move any legacy single contact_date into call slot 1
    if 'call_date_1' not in existing:
        conn.execute("""UPDATE customers SET call_date_1=contact_date
                        WHERE contact_date IS NOT NULL AND contact_date != ''""")

    # Table to track processed emails by Message-ID
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS processed_emails (
            message_id TEXT PRIMARY KEY,
            processed_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS customer_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            filename TEXT NOT NULL,
            filepath TEXT NOT NULL,
            uploaded_at TEXT NOT NULL,
            FOREIGN KEY (customer_id) REFERENCES customers(id)
        );
        CREATE TABLE IF NOT EXISTS policy_documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER,
            policy_number TEXT,
            filename TEXT NOT NULL,
            filepath TEXT NOT NULL,
            received_at TEXT NOT NULL,
            message_id TEXT UNIQUE,
            whatsapp_sent_at TEXT,
            FOREIGN KEY (customer_id) REFERENCES customers(id)
        );
        CREATE TABLE IF NOT EXISTS policy_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            policy_document_id INTEGER,
            customer_id INTEGER,
            policy_number TEXT,
            doc_type_label TEXT,
            doc_type_code TEXT,
            branch TEXT,
            agent_name TEXT,
            agent_number TEXT,
            insured_name TEXT,
            insured_id TEXT,
            spouse_id TEXT,
            address TEXT,
            phone_mobile TEXT,
            phone_home TEXT,
            email TEXT,
            period_start TEXT,
            period_end TEXT,
            premium TEXT,
            total_payment TEXT,
            doc_date TEXT,
            extracted_at TEXT NOT NULL,
            FOREIGN KEY (policy_document_id) REFERENCES policy_documents(id),
            FOREIGN KEY (customer_id) REFERENCES customers(id)
        );
        CREATE TABLE IF NOT EXISTS insureds (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_number TEXT UNIQUE,
            name TEXT,
            agency TEXT,
            brand TEXT,
            phone TEXT,
            email TEXT,
            address TEXT,
            policy_number TEXT,
            period_start TEXT,
            period_end TEXT,
            status TEXT DEFAULT 'פעיל',
            status_override INTEGER DEFAULT 0,
            whatsapp_source TEXT,
            agent_notes TEXT,
            is_vip INTEGER DEFAULT 0,
            handled_by TEXT,
            call_date_1 TEXT, call_status_1 TEXT, call_by_1 TEXT,
            call_date_2 TEXT, call_status_2 TEXT, call_by_2 TEXT,
            call_date_3 TEXT, call_status_3 TEXT, call_by_3 TEXT,
            created_at TEXT,
            updated_at TEXT
        );
        CREATE TABLE IF NOT EXISTS unmatched_submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            received_at TEXT,
            subject TEXT,
            name TEXT, id_number TEXT, phone TEXT, email TEXT,
            brand TEXT, installments TEXT, payment_method TEXT,
            card_number TEXT, card_expiry TEXT, card_holder_id TEXT,
            coverage TEXT, comments TEXT,
            status TEXT DEFAULT 'pending',
            admin_note TEXT,
            message_id TEXT UNIQUE
        );
    ''')

    # Add card + tracking columns if missing
    existing = [r[1] for r in conn.execute("PRAGMA table_info(customers)").fetchall()]
    for col, typ in [('form_card_number','TEXT'), ('form_card_expiry','TEXT'),
                     ('form_id_card_holder','TEXT'), ('handled_by','TEXT')]:
        if col not in existing:
            conn.execute(f"ALTER TABLE customers ADD COLUMN {col} {typ}")

    # Add handled_by to unmatched_submissions if missing
    existing_us = [r[1] for r in conn.execute("PRAGMA table_info(unmatched_submissions)").fetchall()]
    if 'handled_by' not in existing_us:
        conn.execute("ALTER TABLE unmatched_submissions ADD COLUMN handled_by TEXT")
    # One-time cleanup: purge automated morning monitor tests captured before the
    # ingestion-level filter existed. Marker-based rows, plus fully-empty rows (the
    # no-field monitor forms like 'מינוי סוכן' leave no identifying data → not actionable).
    conn.execute(
        "DELETE FROM unmatched_submissions WHERE "
        "COALESCE(id_number,'')='999999999' OR COALESCE(email,'')='monitor-check@example.com' "
        "OR COALESCE(name,'')='MONITOR-CHECK-DO-NOT-PROCESS' OR ("
        "COALESCE(name,'')='' AND COALESCE(id_number,'')='' "
        "AND COALESCE(phone,'')='' AND COALESCE(email,'')='')"
    )
    # Add doc_date to policy_records if missing; backfill from linked document date
    existing_pr = [r[1] for r in conn.execute("PRAGMA table_info(policy_records)").fetchall()]
    if 'doc_date' not in existing_pr:
        conn.execute("ALTER TABLE policy_records ADD COLUMN doc_date TEXT")
        conn.execute("""UPDATE policy_records SET doc_date=(
            SELECT received_at FROM policy_documents WHERE policy_documents.id=policy_records.policy_document_id)
            WHERE doc_date IS NULL""")
    conn.commit()

    # Zero-pad short numeric ID numbers to 9 digits (idempotent — once padded,
    # length is 9 so the row is no longer selected).
    short_ids = conn.execute(
        "SELECT id, id_number FROM customers "
        "WHERE id_number GLOB '[0-9]*' AND id_number NOT GLOB '*[^0-9]*' AND length(id_number) < 9"
    ).fetchall()
    for row in short_ids:
        conn.execute("UPDATE customers SET id_number=? WHERE id=?",
                     (row[1].zfill(9), row[0]))
    if short_ids:
        conn.commit()

    # Default admin
    if not conn.execute("SELECT id FROM users WHERE username='sharon'").fetchone():
        conn.execute(
            "INSERT INTO users (username, password_hash, display_name, role) VALUES (?,?,?,?)",
            ('sharon', generate_password_hash('admin123'), 'שרון', 'admin')
        )
    conn.commit()

    # First-time backfill of the insureds master from existing policy PDFs
    have_insureds = conn.execute("SELECT COUNT(*) FROM insureds").fetchone()[0]
    have_records = conn.execute("SELECT COUNT(*) FROM policy_records").fetchone()[0]
    if have_insureds == 0 and have_records > 0:
        try:
            rebuild_insureds(conn)
        except Exception as e:
            print(f'[init] insureds backfill שגיאה: {e}')

    conn.close()

def active_month():
    conn = get_db()
    m = conn.execute("SELECT * FROM months WHERE is_active=1 ORDER BY id DESC LIMIT 1").fetchone()
    conn.close()
    return m

# ── Auth ────────────────────────────────────────────────────

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if session.get('role') != 'admin':
            flash('גישה מנהל בלבד', 'danger')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated

# ── Routes ──────────────────────────────────────────────────

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username'].strip()
        password = request.form['password']
        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        conn.close()
        if user and check_password_hash(user['password_hash'], password):
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['display_name'] = user['display_name']
            session['role'] = user['role']
            return redirect(url_for('index'))
        flash('שם משתמש או סיסמה שגויים', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    month = active_month()
    stats = {}
    if month:
        conn = get_db()
        rows = conn.execute("""SELECT status, brand, form_received_at,
                               call_status_1, call_status_2, call_status_3
                               FROM customers WHERE month_id=?""", (month['id'],)).fetchall()
        total = len(rows)
        renewed = sum(1 for r in rows if r['status'] == 'חודש')
        no_renew = sum(1 for r in rows if r['status'] == 'לא רוצים לחדש')
        seen = sum(1 for r in rows if r['status'] == 'לקוח ענה/ V כחול')
        forms = sum(1 for r in rows if r['status'] == 'טופס התקבל')
        renewed_from_forms = sum(1 for r in rows if r['status'] == 'חודש' and r['form_received_at'])
        pending = total - renewed - no_renew - seen - forms
        # "No contact made" — still pending (blank status) AND no call attempt logged.
        # A customer with any אין מענה 1/2/3 counts once as contacted and drops out.
        def _contacted(r):
            return bool(r['call_status_1'] or r['call_status_2'] or r['call_status_3'])
        no_contact = sum(1 for r in rows if not r['status'] and not _contacted(r))
        gaia = sum(1 for r in rows if r['brand'] == 'גאיה')
        winner = sum(1 for r in rows if r['brand'] in ('ווינר', 'אופיר'))
        gaia_renewed = sum(1 for r in rows if r['brand'] == 'גאיה' and r['status'] == 'חודש')
        winner_renewed = sum(1 for r in rows if r['brand'] in ('ווינר', 'אופיר') and r['status'] == 'חודש')
        # 'pending' = items a rep explicitly escalated to the admin queue (mark_clarify).
        # Raw website intake stays 'ממתין' and lives in /admin/other-forms, not here.
        unmatched = conn.execute("SELECT COUNT(*) FROM unmatched_submissions WHERE status='pending'").fetchone()[0]
        conn.close()
        stats = dict(total=total, renewed=renewed, no_renew=no_renew, seen=seen, forms=forms, pending=pending,
                     renewed_from_forms=renewed_from_forms, no_contact=no_contact,
                     gaia=gaia, winner=winner, gaia_renewed=gaia_renewed, winner_renewed=winner_renewed,
                     pct=round(renewed / total * 100, 1) if total else 0, unmatched=unmatched)
    return render_template('dashboard.html', month=month, stats=stats)

@app.route('/customers')
@login_required
def customers():
    month = active_month()
    if not month:
        flash('אין חודש פעיל. המנהל צריך לטעון נתונים.', 'warning')
        return redirect(url_for('index'))

    brand_filter = request.args.get('brand', '')
    status_filter = request.args.get('status', '')
    search = request.args.get('q', '').strip()

    query = "SELECT * FROM customers WHERE month_id=?"
    params = [month['id']]

    if brand_filter:
        if brand_filter == 'ווינר':
            query += " AND brand IN ('ווינר','אופיר')"
        else:
            query += " AND brand=?"
            params.append(brand_filter)
    if status_filter == '__empty__':
        query += " AND (status IS NULL OR status='')"
    elif status_filter:
        query += " AND status=?"
        params.append(status_filter)
    if search:
        query += " AND (name LIKE ? OR phone LIKE ? OR policy_number LIKE ?)"
        like = f'%{search}%'
        params += [like, like, like]

    query += " ORDER BY name"

    conn = get_db()
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return render_template('customers.html', customers=rows, month=month,
                           brand_filter=brand_filter, status_filter=status_filter, search=search,
                           statuses=STATUSES)

@app.route('/search')
@login_required
def search_customers():
    """Global customer search (across all months) by name, phone or policy number."""
    search = request.args.get('q', '').strip()
    rows = []
    if search:
        conn = get_db()
        like = f'%{search}%'
        # Normalised phone match too, so 050-123 finds 0501234567 etc.
        digits = re.sub(r'\D', '', search)
        phone_like = f'%{digits}%' if digits else like
        rows = conn.execute(
            """SELECT c.*, m.name AS month_name
               FROM customers c
               LEFT JOIN months m ON m.id = c.month_id
               WHERE c.name LIKE ?
                  OR c.phone LIKE ?
                  OR replace(replace(c.phone,'-',''),' ','') LIKE ?
                  OR c.policy_number LIKE ?
                  OR ltrim(c.id_number,'0') LIKE ?
               ORDER BY m.id DESC, c.name""",
            (like, like, phone_like, like, like)
        ).fetchall()
        conn.close()
    return render_template('search_results.html', customers=rows, search=search)


@app.route('/customer/<int:cid>', methods=['GET', 'POST'])
@login_required
def customer_detail(cid):
    conn = get_db()
    month = active_month()
    customer = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    conn.close()
    if not customer:
        flash('לקוח לא נמצא', 'danger')
        return redirect(url_for('customers'))
    wa_link = build_followup_wa_link(customer)
    return render_template('customer_detail.html', c=customer, month=month,
                           statuses=STATUSES, wa_link=wa_link)


def build_followup_wa_link(customer):
    """Pre-filled WhatsApp reminder link for a customer who didn't answer calls."""
    from urllib.parse import quote
    phone = re.sub(r'\D', '', str(customer['phone'] or ''))
    if not phone:
        return None
    if phone.startswith('0'):
        phone = phone[1:]
    phone = '972' + phone
    site = 'https://www.winner-ins.co.il/renew' if customer['brand'] in ('ווינר', 'אופיר') \
        else 'https://www.gaia-ins.co.il/renew'
    msg = ('היי, \nניסינו להשיג אותך לחידוש הפוליסה. נשמח אם תוכל ליצור איתנו קשר '
           'לטובת החידוש, או לחדש את הפוליסה אונליין באתר ' + site)
    return f'https://wa.me/{phone}?text={quote(msg)}'

@app.route('/customer/<int:cid>/update', methods=['POST'])
@login_required
def update_customer(cid):
    data = request.json or {}
    allowed = ['status', 'agent_notes', 'contact_date', 'interested_in_products',
                'whatsapp_sent_date', 'sharon_notes', 'requests_to_sharon', 'is_vip',
                'whatsapp_source', 'brand',
                'call_date_1', 'call_status_1', 'call_by_1',
                'call_date_2', 'call_status_2', 'call_by_2',
                'call_date_3', 'call_status_3', 'call_by_3']
    # Agents cannot update sharon fields or brand (admin-only)
    if session.get('role') != 'admin':
        for f in ['sharon_notes', 'requests_to_sharon', 'brand']:
            data.pop(f, None)

    # Ofir customers are contacted from Winner's WhatsApp number
    if data.get('brand') == 'אופיר' and 'whatsapp_source' not in data:
        data['whatsapp_source'] = 'ווינר'

    agent = session.get('display_name') or session.get('username', '')
    conn = get_db()

    # Auto-capture the rep who logged a call attempt (like the date) — only when
    # that attempt's date is newly set or changed, so it isn't reassigned on every save.
    if agent and any(f'call_date_{n}' in data for n in (1, 2, 3)):
        prev = conn.execute(
            "SELECT call_date_1, call_date_2, call_date_3 FROM customers WHERE id=?", (cid,)
        ).fetchone()
        for n in (1, 2, 3):
            key = f'call_date_{n}'
            if key in data and data[key] and (not prev or data[key] != prev[f'call_date_{n}']):
                data[f'call_by_{n}'] = agent

    sets = ', '.join(f"{k}=?" for k in data if k in allowed)
    vals = [data[k] for k in data if k in allowed]
    if not sets:
        conn.close()
        return jsonify({'ok': False})
    # Track who changed the status
    if 'status' in data and agent:
        sets += ', handled_by=?'
        vals.append(agent)
    vals.append(cid)
    conn.execute(f"UPDATE customers SET {sets} WHERE id=?", vals)
    conn.commit()
    conn.close()
    return jsonify({'ok': True})

# ── Admin ───────────────────────────────────────────────────

@app.route('/customer/<int:cid>/delete', methods=['POST'])
@login_required
@admin_required
def delete_customer(cid):
    conn = get_db()
    conn.execute("DELETE FROM customers WHERE id=?", (cid,))
    conn.execute("DELETE FROM customer_attachments WHERE customer_id=?", (cid,))
    conn.commit()
    conn.close()
    flash('לקוח נמחק', 'warning')
    return redirect(url_for('customers'))


@app.route('/export/customers-excel')
@login_required
@admin_required
def export_customers_excel():
    import openpyxl
    from io import BytesIO
    conn = get_db()
    month = active_month()
    if not month:
        flash('אין חודש פעיל', 'warning')
        return redirect(url_for('customers'))
    rows = conn.execute("SELECT * FROM customers WHERE month_id=? ORDER BY id", (month['id'],)).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = month['name']

    headers = ['פוליסה', 'שם', 'ת.ז', 'טלפון', 'מותג', 'סטטוס',
               'פרמיה שנה שעברה', 'וואטסאפ נשלח', 'תאריך התקשרות',
               'הערות נציג', 'הערות שרון', 'בקשות משרון',
               'טופס התקבל', 'מייל לקוח', 'תשלומים', 'גבייה',
               'מספר כרטיס', 'תוקף כרטיס', 'הערות טופס', 'טיפל']
    ws.append(headers)

    for r in rows:
        ws.append([
            r['policy_number'], r['name'], r['id_number'], r['phone'], r['brand'], r['status'],
            r['premium_last_year'], r['whatsapp_sent_date'], r['contact_date'],
            r['agent_notes'], r['sharon_notes'], r['requests_to_sharon'],
            r['form_received_at'], r['form_email'], r['form_installments'], r['form_payment_method'],
            r['form_card_number'], r['form_card_expiry'], r['form_comments'], r['handled_by'],
        ])

    # Auto-width
    for col in ws.columns:
        max_len = max((len(str(c.value or '')) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"לקוחות_{month['name']}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin-queue')
@login_required
@admin_required
def admin_queue():
    conn = get_db()
    # Only rep-escalated items ('pending', set by mark_clarify). Raw website intake
    # ('ממתין') belongs to /admin/other-forms — keeping them apart avoids duplication.
    items = conn.execute(
        "SELECT * FROM unmatched_submissions WHERE status='pending' ORDER BY received_at DESC"
    ).fetchall()
    conn.close()
    return render_template('admin_queue.html', items=items)

def guess_category(subject, source):
    """Rough auto-tag for the 'other forms' catch-all — a hint, not a strict classifier."""
    text = subject or ''
    if any(k in text for k in ['כרטיס אשראי', 'אשראי', 'עדכון פרטי תשלום', 'שינוי אמצעי']):
        return 'עדכון אמצעי תשלום'
    if source == 'policy':
        return 'פוליסה לא משויכת (עסקה חדשה?)'
    if 'חדש' in text:
        return 'הצעה חדשה'
    return 'אחר'

@app.route('/admin/other-forms')
@login_required
@admin_required
def other_forms():
    """Catch-all view: every incoming email that didn't become a matched renewal —
    backup net + light organization, regardless of source table."""
    conn = get_db()
    rows = []

    # Only real website-form submissions here. Harel policy PDFs are intentionally
    # excluded — they already live in the insureds master ("כל הלקוחות"), so showing
    # them here too was double-bookkeeping. Automated morning monitor tests are filtered.
    for r in conn.execute(
        "SELECT * FROM unmatched_submissions WHERE status='ממתין' "
        "AND COALESCE(id_number,'') != '999999999' "
        "AND COALESCE(email,'') != 'monitor-check@example.com' "
        "AND COALESCE(name,'') != 'MONITOR-CHECK-DO-NOT-PROCESS' "
        "ORDER BY received_at DESC"
    ).fetchall():
        d = dict(r)
        rows.append({
            'id': d['id'], 'received_at': d['received_at'], 'subject': d['subject'],
            'title': d['name'] or '(ללא שם)', 'detail': d['id_number'] or d['phone'] or '',
            'source': 'טופס', 'category': guess_category(d['subject'], 'form'),
            'link': None, 'kind': 'form', 'full': d,
        })

    rows.sort(key=lambda x: x['received_at'] or '', reverse=True)
    conn.close()
    return render_template('other_forms.html', items=rows)

@app.route('/admin/other-forms/delete', methods=['POST'])
@login_required
@admin_required
def other_forms_delete():
    """Bulk-delete selected rows from the other-forms catch-all (form or policy items)."""
    selected = request.form.getlist('selected')
    form_ids = [s.split(':', 1)[1] for s in selected if s.startswith('form:')]
    policy_ids = [s.split(':', 1)[1] for s in selected if s.startswith('policy:')]

    conn = get_db()
    if form_ids:
        placeholders = ','.join('?' * len(form_ids))
        conn.execute(f"DELETE FROM unmatched_submissions WHERE id IN ({placeholders})", form_ids)
    if policy_ids:
        placeholders = ','.join('?' * len(policy_ids))
        conn.execute(f"DELETE FROM policy_documents WHERE id IN ({placeholders})", policy_ids)
    conn.commit()
    conn.close()
    flash(f'{len(form_ids) + len(policy_ids)} פריטים נמחקו', 'success')
    return redirect(url_for('other_forms'))

@app.route('/admin/policy-records')
@login_required
@admin_required
def policy_records():
    """All customers (master) — one row per insured (by ID), built from the Harel
    policy PDFs. Best-effort extraction — some fields may need correction."""
    q = request.args.get('q', '').strip()
    conn = get_db()
    recompute_insured_statuses(conn)  # keep פעיל/לא פעיל current on view
    if q:
        like = f'%{q}%'
        rows = conn.execute(
            '''SELECT * FROM insureds
               WHERE name LIKE ? OR id_number LIKE ? OR policy_number LIKE ?
                  OR phone LIKE ? OR email LIKE ?
               ORDER BY name''',
            (like, like, like, like, like)
        ).fetchall()
    else:
        rows = conn.execute('SELECT * FROM insureds ORDER BY name LIMIT 500').fetchall()
    total = conn.execute('SELECT COUNT(*) FROM insureds').fetchone()[0]
    conn.close()
    return render_template('policy_records.html', items=rows, q=q, total=total,
                           backfill=_backfill_state)

def build_followup_wa_link_generic(phone, brand):
    """Pre-filled WhatsApp reminder link from a phone + brand (works for insureds too)."""
    from urllib.parse import quote
    p = re.sub(r'\D', '', str(phone or ''))
    if not p:
        return None
    if p.startswith('0'):
        p = p[1:]
    p = '972' + p
    site = 'https://www.winner-ins.co.il/renew' if brand in ('ווינר', 'אופיר') \
        else 'https://www.gaia-ins.co.il/renew'
    msg = ('היי, \nניסינו להשיג אותך לחידוש הפוליסה. נשמח אם תוכל ליצור איתנו קשר '
           'לטובת החידוש, או לחדש את הפוליסה אונליין באתר ' + site)
    return f'https://wa.me/{p}?text={quote(msg)}'

@app.route('/insured/<int:iid>')
@login_required
@admin_required
def insured_detail(iid):
    conn = get_db()
    ins = conn.execute("SELECT * FROM insureds WHERE id=?", (iid,)).fetchone()
    if not ins:
        conn.close()
        flash('לקוח לא נמצא', 'danger')
        return redirect(url_for('policy_records'))
    # PDF history for this insured (by ID), newest policy first
    docs = conn.execute(
        """SELECT pd.id AS doc_id, pd.filename, pd.received_at,
                  pr.doc_type_label, pr.period_start, pr.period_end
           FROM policy_records pr JOIN policy_documents pd ON pr.policy_document_id = pd.id
           WHERE ltrim(pr.insured_id,'0') = ltrim(?,'0')
           ORDER BY pr.extracted_at DESC""",
        (ins['id_number'],)
    ).fetchall()
    conn.close()
    wa_link = build_followup_wa_link_generic(ins['phone'], ins['brand'])
    return render_template('insured_detail.html', c=ins, docs=docs, wa_link=wa_link)

@app.route('/insured/<int:iid>/update', methods=['POST'])
@login_required
@admin_required
def insured_update(iid):
    data = request.json or {}
    allowed = ['agent_notes', 'whatsapp_source', 'is_vip',
               'call_date_1', 'call_status_1', 'call_by_1',
               'call_date_2', 'call_status_2', 'call_by_2',
               'call_date_3', 'call_status_3', 'call_by_3']
    agent = session.get('display_name') or session.get('username', '')
    conn = get_db()

    # Manual status change is an admin override that sticks (req 8)
    if 'status' in data and data['status']:
        conn.execute("UPDATE insureds SET status=?, status_override=1, updated_at=? WHERE id=?",
                     (data['status'], datetime.datetime.now().isoformat(), iid))

    # Auto-capture the rep who logged a call attempt (like the renewals page)
    if agent and any(f'call_date_{n}' in data for n in (1, 2, 3)):
        prev = conn.execute(
            "SELECT call_date_1, call_date_2, call_date_3 FROM insureds WHERE id=?", (iid,)
        ).fetchone()
        for n in (1, 2, 3):
            key = f'call_date_{n}'
            if key in data and data[key] and (not prev or data[key] != prev[f'call_date_{n}']):
                data[f'call_by_{n}'] = agent

    sets = ', '.join(f"{k}=?" for k in data if k in allowed)
    if sets:
        vals = [data[k] for k in data if k in allowed]
        vals.append(iid)
        conn.execute(f"UPDATE insureds SET {sets} WHERE id=?", vals)
    conn.commit()
    conn.close()
    return jsonify({'ok': True})

@app.route('/customer/<int:cid>/clarify', methods=['POST'])
@login_required
def mark_clarify(cid):
    """Move customer to admin queue for clarification. Requires a reason (rep notes)."""
    data = request.get_json(silent=True) or {}
    note = (data.get('note') or '').strip()
    if not note:
        return jsonify({'ok': False, 'error': 'נא לפרט את הסיבה להעברה לאדמין'}), 400
    agent = session.get('display_name') or session.get('username', '')
    conn = get_db()
    c = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    if c:
        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
        comment_parts = []
        if c['form_comments']: comment_parts.append(c['form_comments'])
        if note: comment_parts.append(note)
        comments = ' | '.join(comment_parts)
        # Use INSERT OR REPLACE so re-clarifying the same customer works
        conn.execute('''INSERT OR REPLACE INTO unmatched_submissions
            (received_at, subject, name, id_number, phone, email, brand, installments,
             payment_method, card_number, card_expiry, card_holder_id, coverage, comments,
             status, handled_by, message_id)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,'pending',?,?)''',
            (now, 'דורש בירור', c['name'], c['id_number'], c['phone'],
             c['form_email'] or '', c['brand'], c['form_installments'] or '',
             c['form_payment_method'] or '', c['form_card_number'] or '',
             c['form_card_expiry'] or '', c['form_id_card_holder'] or '',
             c['form_coverage'] or '', comments, agent, f'queue-cid-{cid}'))
        conn.execute("UPDATE customers SET status='דורש בירור', handled_by=? WHERE id=?", (agent, cid))
        conn.commit()
    conn.close()
    return jsonify({'ok': True})

@app.route('/admin-queue/<int:sid>/action', methods=['POST'])
@login_required
@admin_required
def admin_queue_action(sid):
    action = request.form.get('action')
    note = request.form.get('admin_note', '')
    conn = get_db()
    if action == 'dismiss':
        conn.execute("UPDATE unmatched_submissions SET status='dismissed', admin_note=? WHERE id=?", (note, sid))
    elif action == 'link':
        cid = request.form.get('customer_id', '')
        if cid:
            sub = conn.execute("SELECT * FROM unmatched_submissions WHERE id=?", (sid,)).fetchone()
            if sub:
                now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
                conn.execute("""UPDATE customers SET status='טופס התקבל',
                    form_email=?, form_installments=?, form_payment_method=?,
                    form_received_at=?, form_coverage=?, form_comments=?,
                    form_card_number=?, form_card_expiry=?, form_id_card_holder=?
                    WHERE id=?""",
                    (sub['email'], sub['installments'], sub['payment_method'], now,
                     sub['coverage'], sub['comments'], sub['card_number'],
                     sub['card_expiry'], sub['card_holder_id'], cid))
                conn.execute("UPDATE unmatched_submissions SET status='linked', admin_note=? WHERE id=?",
                             (f'שויך ללקוח {cid}', sid))
    elif action == 'resolve':
        # For clarify items — set final status on the linked customer
        new_status = request.form.get('new_status', '')
        sub = conn.execute("SELECT * FROM unmatched_submissions WHERE id=?", (sid,)).fetchone()
        if sub and new_status:
            # Extract customer id from message_id = 'queue-cid-{cid}'
            msg_id = sub['message_id'] or ''
            if msg_id.startswith('queue-cid-'):
                cid = msg_id.replace('queue-cid-', '')
                agent = session.get('display_name') or session.get('username', '')
                conn.execute("UPDATE customers SET status=?, handled_by=? WHERE id=?",
                             (new_status, agent, cid))
            conn.execute("UPDATE unmatched_submissions SET status='resolved', admin_note=? WHERE id=?",
                         (f'סטטוס עודכן: {new_status} | {note}', sid))
    conn.commit()
    conn.close()
    flash('בוצע', 'success')
    return redirect(url_for('admin_queue'))

@app.route('/attachment/<int:att_id>')
@login_required
def download_attachment(att_id):
    conn = get_db()
    att = conn.execute('SELECT * FROM customer_attachments WHERE id=?', (att_id,)).fetchone()
    conn.close()
    if not att:
        return 'לא נמצא', 404
    safe_name = re.sub(r'[\r\n]+', ' ', att['filename']).strip()
    return send_file(att['filepath'], as_attachment=True, download_name=safe_name)

@app.route('/policy-document/<int:doc_id>')
@login_required
def download_policy_document(doc_id):
    conn = get_db()
    doc = conn.execute('SELECT * FROM policy_documents WHERE id=?', (doc_id,)).fetchone()
    conn.close()
    if not doc:
        return 'לא נמצא', 404
    safe_name = re.sub(r'[\r\n]+', ' ', doc['filename']).strip()
    return send_file(doc['filepath'], as_attachment=True, download_name=safe_name)


@app.route('/queue')
@login_required
def queue():
    month = active_month()
    if not month:
        flash('אין חודש פעיל', 'warning')
        return redirect(url_for('index'))
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM customers WHERE month_id=? AND status='טופס התקבל' ORDER BY form_received_at DESC",
        (month['id'],)
    ).fetchall()
    # Fetch attachments per customer
    attachments = {}
    for r in rows:
        atts = conn.execute(
            'SELECT * FROM customer_attachments WHERE customer_id=?', (r['id'],)
        ).fetchall()
        if atts:
            attachments[r['id']] = atts
    conn.close()
    return render_template('queue.html', customers=rows, month=month, attachments=attachments)


@app.route('/admin')
@login_required
@admin_required
def admin():
    conn = get_db()
    users = conn.execute("SELECT id, username, display_name, role FROM users ORDER BY role, display_name").fetchall()
    months = conn.execute("SELECT * FROM months ORDER BY id DESC").fetchall()
    conn.close()
    return render_template('admin.html', users=users, months=months,
                           email_sync_enabled=EMAIL_CONFIG['enabled'])

@app.route('/admin/import', methods=['POST'])
@login_required
@admin_required
def import_excel():
    f = request.files.get('file')
    month_name = request.form.get('month_name', '').strip()
    if not f or not month_name:
        flash('חסר קובץ או שם חודש', 'danger')
        return redirect(url_for('admin'))

    try:
        wb = load_workbook(f, data_only=True)
        ws = wb.active

        conn = get_db()
        # Req 4+5: before starting the new cycle, move the outgoing month's customers
        # into the "all customers" master (renewed → פעיל, else → לא פעיל), preserving activity.
        prev = conn.execute("SELECT id FROM months WHERE is_active=1 ORDER BY id DESC LIMIT 1").fetchone()
        promoted = 0
        if prev:
            promoted = promote_customers_to_insureds(conn, prev['id'])
        # Deactivate all months
        conn.execute("UPDATE months SET is_active=0")
        # Create new month
        conn.execute("INSERT INTO months (name, created_at, is_active) VALUES (?,?,1)",
                     (month_name, datetime.datetime.now().isoformat()))
        month_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]

        count = 0
        # Find header row (row with 'פוליסה')
        header_row = None
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if row and 'פוליסה' in str(row[0]):
                header_row = i
                break

        if not header_row:
            # Fallback: assume row 3
            header_row = 3

        headers = [str(c).strip() if c else '' for c in list(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))[0]]

        def col(name, row_vals):
            try:
                idx = next(i for i, h in enumerate(headers) if name in h)
                return str(row_vals[idx]).strip() if row_vals[idx] is not None else ''
            except StopIteration:
                return ''

        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if not row or not row[0]:
                continue
            policy = str(row[0]).strip() if row[0] else ''
            if not policy or policy in ('None', ''):
                continue
            # Always increment policy number by 1 (each year the last digit advances: 5→6, 6→7, etc.)
            if policy.isdigit():
                policy = str(int(policy) + 1)
            name = col('שם', row)
            if not name or name == 'None':
                continue
            raw_status = col('סטטוס', row)
            status_map = {
                'חודש': 'חודש',
                'לא חודש': '',
                'לא רוצים לחדש': 'לא רוצים לחדש',
                'לקוח ענה/ V כחול': 'לקוח ענה/ V כחול',
            }
            mapped_status = status_map.get(raw_status, '')

            row_brand = col('מותג', row)
            conn.execute("""
                INSERT INTO customers
                (month_id, policy_number, name, id_number, phone, brand, status,
                 premium_last_year, whatsapp_sent_date, sharon_notes, requests_to_sharon,
                 contact_date, agent_notes, interested_in_products, whatsapp_source)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                month_id, policy, name,
                normalize_id_number(col('ת.ז', row)), col('טלפון', row), row_brand, mapped_status,
                col('פרמיה', row), col('וואטסאפ', row), col('הערות שרון', row),
                col('בקשות משרון', row), col('תאריך התקשרות', row),
                col('הערות חידושים', row), col('מתעניין', row),
                'ווינר' if row_brand == 'אופיר' else None
            ))
            count += 1

        conn.commit()
        conn.close()
        msg = f'נטענו {count} לקוחות לחודש {month_name}'
        if promoted:
            msg += f' · {promoted} לקוחות מהחודש הקודם עברו ל"כל הלקוחות"'
        flash(msg, 'success')
    except Exception as e:
        flash(f'שגיאה בייבוא: {e}', 'danger')

    return redirect(url_for('admin'))

@app.route('/admin/users/add', methods=['POST'])
@login_required
@admin_required
def add_user():
    username = request.form['username'].strip()
    display_name = request.form['display_name'].strip()
    password = request.form['password']
    role = request.form.get('role', 'agent')
    try:
        conn = get_db()
        conn.execute("INSERT INTO users (username, password_hash, display_name, role) VALUES (?,?,?,?)",
                     (username, generate_password_hash(password), display_name, role))
        conn.commit()
        conn.close()
        flash(f'משתמש {display_name} נוצר', 'success')
    except Exception as e:
        flash(f'שגיאה: {e}', 'danger')
    return redirect(url_for('admin'))

@app.route('/admin/users/delete/<int:uid>', methods=['POST'])
@login_required
@admin_required
def delete_user(uid):
    if uid == session['user_id']:
        flash('לא ניתן למחוק את עצמך', 'danger')
        return redirect(url_for('admin'))
    conn = get_db()
    conn.execute("DELETE FROM users WHERE id=?", (uid,))
    conn.commit()
    conn.close()
    flash('משתמש נמחק', 'success')
    return redirect(url_for('admin'))

@app.route('/admin/users/reset-password/<int:uid>', methods=['POST'])
@login_required
@admin_required
def reset_password(uid):
    new_pass = request.form['new_password']
    conn = get_db()
    conn.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(new_pass), uid))
    conn.commit()
    conn.close()
    flash('סיסמה עודכנה', 'success')
    return redirect(url_for('admin'))

@app.route('/export/wasender')
@login_required
@admin_required
def export_wasender():
    month = active_month()
    if not month:
        flash('אין חודש פעיל', 'danger')
        return redirect(url_for('index'))

    brand_filter = request.args.get('brand', '')
    mark_sent = request.args.get('mark_sent', '0') == '1'
    # mode: 'first' = all without whatsapp sent | 'reminder' = didn't renew and don't want to cancel
    mode = request.args.get('mode', 'first')

    conn = get_db()

    if mode == 'first':
        # First send: everyone who hasn't received WhatsApp yet
        query = """SELECT id, name, phone FROM customers
                   WHERE month_id=? AND (whatsapp_sent_date IS NULL OR whatsapp_sent_date='')"""
    else:
        # Reminder: only those who haven't renewed and didn't say they don't want to renew
        query = """SELECT id, name, phone FROM customers
                   WHERE month_id=? AND (status IS NULL OR status='' OR status='לקוח ענה/ V כחול')"""

    params = [month['id']]
    if brand_filter:
        if brand_filter == 'ווינר':
            query += " AND brand IN ('ווינר','אופיר')"
        else:
            query += " AND brand=?"
            params.append(brand_filter)
    query += " ORDER BY name"

    rows = conn.execute(query, params).fetchall()

    wb = NewWorkbook()
    ws = wb.active
    ws.title = 'WASender'
    ws.append(['phone', 'name'])

    today = datetime.date.today().isoformat()
    ids = []
    for r in rows:
        phone = str(r['phone']).replace('-', '').replace(' ', '')
        if phone.startswith('0'):
            phone = '972' + phone[1:]
        ws.append([phone, r['name']])
        ids.append(r['id'])

    if mark_sent and ids:
        placeholders = ','.join('?' * len(ids))
        conn.execute(f"UPDATE customers SET whatsapp_sent_date=? WHERE id IN ({placeholders})",
                     [today] + ids)
        conn.commit()

    conn.close()

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    mode_label = 'ראשונה' if mode == 'first' else 'תזכורת'
    filename = f"wasender_{mode_label}_{month['name'].replace(' ','_')}_{today}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/renewal', methods=['POST'])
def api_renewal():
    """Receives form submissions from winner-ins.co.il/renew and gaia-ins.co.il/renew"""
    data = request.json or request.form.to_dict()

    id_number = normalize_id_number(data.get('id_number') or data.get('id'))
    phone = str(data.get('phone') or data.get('telephone') or '').strip()
    name = str(data.get('name') or data.get('full_name') or '').strip()
    email = str(data.get('email') or '').strip()
    installments = str(data.get('installments') or data.get('payment_installments') or '').strip()
    payment_method = str(data.get('payment_method') or '').strip()
    comments = str(data.get('comments') or '').strip()
    brand = str(data.get('brand') or '').strip()

    if not id_number and not phone:
        return jsonify({'ok': False, 'error': 'missing id or phone'}), 400

    month = active_month()
    if not month:
        return jsonify({'ok': False, 'error': 'no active month'}), 400

    conn = get_db()
    customer = None
    if id_number:
        norm_id = id_number.lstrip('0')
        customer = conn.execute(
            "SELECT * FROM customers WHERE month_id=? AND ltrim(id_number,'0')=?",
            (month['id'], norm_id)
        ).fetchone()
    if not customer and phone:
        clean_phone = phone.replace('-', '').replace(' ', '')
        customer = conn.execute(
            "SELECT * FROM customers WHERE month_id=? AND replace(replace(phone,'-',''),' ','')=?",
            (month['id'], clean_phone)
        ).fetchone()

    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')

    if customer:
        conn.execute("""UPDATE customers SET status='טופס התקבל',
                        form_email=?, form_installments=?, form_payment_method=?,
                        form_received_at=?, form_comments=?
                        WHERE id=?""",
                     (email, installments, payment_method, now, comments, customer['id']))
        conn.commit()
        conn.close()
        return jsonify({'ok': True, 'matched': True, 'customer': customer['name']})
    else:
        conn.execute("""INSERT INTO customers
            (month_id, name, id_number, phone, brand, status,
             form_email, form_installments, form_payment_method, form_received_at, form_comments,
             whatsapp_source)
            VALUES (?,?,?,?,?,'טופס התקבל',?,?,?,?,?,?)""",
            (month['id'], name, id_number, phone, brand,
             email, installments, payment_method, now, comments,
             'ווינר' if brand == 'אופיר' else None))
        conn.commit()
        conn.close()
        return jsonify({'ok': True, 'matched': False, 'note': 'added as new'})


# ── Email parsing helpers ────────────────────────────────────

def decode_str(s):
    """Decode MIME-encoded email header string."""
    parts = decode_header(s)
    result = ''
    for b, enc in parts:
        if isinstance(b, bytes):
            result += b.decode(enc or 'utf-8', errors='replace')
        else:
            result += b
    # MIME header folding can leave embedded \r\n — breaks HTTP headers (Content-Disposition) if left in
    return re.sub(r'[\r\n]+', ' ', result).strip()

def parse_renewal_email(msg_text, subject=''):
    """
    Parse form fields from renewal emails.
    Format: fields and values are space-separated in sequence (no colons).
    e.g. 'שם מלא ארנה אדם מספר ת.ז 056062608 אימייל ...'
    """
    # Known field tokens in priority order — longer/multi-word first
    FIELDS = [
        'שם מלא', 'מספר ת.ז', 'birth_date', 'אימייל', 'טלפון',
        'coverage_option', 'מספר תשלומים', 'מספר פוליסה',
        'אמצעי גביה', 'מספר כרטיס', 'תוקף כרטיס',
        'ת.ז בעל הכרטיס', 'שם בעל הכרטיס', 'הכרטיס על שם המבוטח',
        'מקצועות נוספים', 'הערות',
    ]

    # Build regex that splits on any known field name
    escaped = [re.escape(f) for f in FIELDS]
    splitter = '(' + '|'.join(escaped) + ')'
    parts = re.split(splitter, msg_text)

    result = {}
    i = 1
    while i < len(parts) - 1:
        key = parts[i].strip()
        val = parts[i + 1].strip() if i + 1 < len(parts) else ''
        # Remove leading/trailing em-dash placeholder
        val = val.strip('— ').strip()
        result[key] = val
        i += 2

    # Brand from subject line: "גאיה | ..." or "ווינר | ..."
    brand = ''
    if 'גאיה' in subject:
        brand = 'גאיה'
    elif 'ווינר' in subject:
        brand = 'ווינר'

    return {
        'name': result.get('שם מלא', ''),
        'id_number': result.get('מספר ת.ז', ''),
        'phone': result.get('טלפון', ''),
        'email': result.get('אימייל', ''),
        'installments': result.get('מספר תשלומים', ''),
        'payment_method': result.get('אמצעי גביה', ''),
        'comments': result.get('הערות', ''),
        'brand': brand,
        'policy_number': result.get('מספר פוליסה', ''),
        'card_number': result.get('מספר כרטיס', ''),
        'card_expiry': result.get('תוקף כרטיס', ''),
        'card_holder_id': result.get('ת.ז בעל הכרטיס', ''),
        'coverage_option': result.get('coverage_option', ''),
    }

def get_email_body(msg):
    """Extract plain text body from an email message."""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get('Content-Disposition', ''))
            if ct == 'text/plain' and 'attachment' not in cd:
                charset = part.get_content_charset() or 'utf-8'
                return part.get_payload(decode=True).decode(charset, errors='replace')
        # Fallback: try HTML
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                charset = part.get_content_charset() or 'utf-8'
                html = part.get_payload(decode=True).decode(charset, errors='replace')
                return re.sub(r'<[^>]+>', ' ', html)
    else:
        charset = msg.get_content_charset() or 'utf-8'
        return msg.get_payload(decode=True).decode(charset, errors='replace')
    return ''

def process_renewal_data(data, message_id='', subject='', received_at=''):
    """
    Match email form data to a customer in the active month.
    - Matched → update customer, status='טופס התקבל', return customer_id
    - Not matched → save to unmatched_submissions for admin review, return None
    """
    id_number      = normalize_id_number(data.get('id_number'))
    phone          = str(data.get('phone') or '').strip()
    name           = str(data.get('name') or '').strip()
    email_val      = str(data.get('email') or '').strip()
    installments   = str(data.get('installments') or '').strip()
    payment_method = str(data.get('payment_method') or '').strip()
    comments       = str(data.get('comments') or '').strip()
    brand          = str(data.get('brand') or '').strip()
    coverage       = str(data.get('coverage_option') or '').strip()
    card_number    = str(data.get('card_number') or '').strip()
    card_expiry    = str(data.get('card_expiry') or '').strip()
    card_holder_id = str(data.get('card_holder_id') or '').strip()

    now = received_at or datetime.datetime.now().strftime('%Y-%m-%d %H:%M')

    conn = get_db()
    month = conn.execute("SELECT * FROM months WHERE is_active=1 ORDER BY id DESC LIMIT 1").fetchone()

    if not month:
        conn.close()
        print('[email-sync] אין חודש פעיל')
        return None

    if not id_number and not phone:
        # No identifying info — send to admin
        conn.execute('''INSERT OR IGNORE INTO unmatched_submissions
            (received_at, subject, name, id_number, phone, email, brand, installments,
             payment_method, card_number, card_expiry, card_holder_id, coverage, comments, message_id)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (now, subject, name, id_number, phone, email_val, brand, installments,
             payment_method, card_number, card_expiry, card_holder_id, coverage, comments, message_id))
        conn.commit()
        conn.close()
        print('[email-sync] חסר מזהה → unmatched')
        return None

    customer = None
    if id_number:
        customer = conn.execute(
            "SELECT * FROM customers WHERE month_id=? AND ltrim(id_number,'0')=?",
            (month['id'], id_number.lstrip('0'))
        ).fetchone()
    if not customer and phone:
        clean = phone.replace('-', '').replace(' ', '')
        customer = conn.execute(
            "SELECT * FROM customers WHERE month_id=? AND replace(replace(phone,'-',''),' ','')=?",
            (month['id'], clean)
        ).fetchone()

    if customer:
        conn.execute("""UPDATE customers SET status='טופס התקבל',
                        form_email=?, form_installments=?, form_payment_method=?,
                        form_received_at=?, form_coverage=?, form_comments=?,
                        form_card_number=?, form_card_expiry=?, form_id_card_holder=?
                        WHERE id=?""",
                     (email_val, installments, payment_method, now, coverage, comments,
                      card_number, card_expiry, card_holder_id, customer['id']))
        conn.commit()
        cid = customer['id']
        conn.close()
        print(f'[email-sync] עודכן: {customer["name"]} → טופס התקבל')
        return cid
    else:
        # No match in current month → admin queue
        conn.execute('''INSERT OR IGNORE INTO unmatched_submissions
            (received_at, subject, name, id_number, phone, email, brand, installments,
             payment_method, card_number, card_expiry, card_holder_id, coverage, comments, message_id)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (now, subject, name, id_number, phone, email_val, brand, installments,
             payment_method, card_number, card_expiry, card_holder_id, coverage, comments, message_id))
        conn.commit()
        conn.close()
        print(f'[email-sync] לא זוהה: {name} → תור אדמין')
        return None

ATTACHMENTS_DIR = os.environ.get('ATTACHMENTS_DIR', os.path.join(os.path.dirname(__file__), 'attachments')).strip()
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

def _save_attachments(msg, customer_id):
    """Extract and save email attachments, record in DB."""
    saved = []
    for part in msg.walk():
        cd = str(part.get('Content-Disposition', ''))
        if 'attachment' not in cd:
            continue
        raw_fn = part.get_filename()
        if not raw_fn:
            continue
        filename = decode_str(raw_fn)
        data = part.get_payload(decode=True)
        if not data:
            continue
        cust_dir = os.path.join(ATTACHMENTS_DIR, str(customer_id))
        os.makedirs(cust_dir, exist_ok=True)
        # Avoid collisions
        safe_fn = re.sub(r'[\\/*?:"<>|]', '_', filename)
        filepath = os.path.join(cust_dir, safe_fn)
        with open(filepath, 'wb') as f:
            f.write(data)
        conn = get_db()
        conn.execute(
            'INSERT INTO customer_attachments (customer_id, filename, filepath, uploaded_at) VALUES (?,?,?,?)',
            (customer_id, filename, filepath, datetime.datetime.now().isoformat())
        )
        conn.commit()
        conn.close()
        saved.append(filename)
        print(f'[email-sync] קובץ נשמר: {filename}')
    return saved


_email_check_lock = threading.Lock()

def check_email_inbox():
    """Connect to IMAP, process renewal emails not yet seen (tracked by Message-ID in DB)."""
    if not _email_check_lock.acquire(blocking=False):
        print('[email-sync] בדיקה כבר רצה — דילוג')
        return 0
    try:
        return _check_email_inbox_impl()
    finally:
        _email_check_lock.release()

def _check_email_inbox_impl():
    cfg = EMAIL_CONFIG
    if not cfg['enabled'] or not cfg['imap_server'] or not cfg['password']:
        return 0

    processed = 0
    try:
        mail = imaplib.IMAP4_SSL(cfg['imap_server'], cfg['imap_port'])
        mail.login(cfg['username'], cfg['password'])
        mail.select('INBOX')

        # Search from Resend since 30 days ago (limits scan size; processed_emails prevents duplicates)
        since_date = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime('%d-%b-%Y')
        status, data = mail.search(None, f'FROM "{cfg["sender_filter"]}" SINCE {since_date}')
        if status != 'OK':
            mail.logout()
            return 0

        conn = get_db()
        # Get month load time to filter only emails after that point
        month = conn.execute("SELECT created_at FROM months WHERE is_active=1 ORDER BY id DESC LIMIT 1").fetchone()
        month_loaded_at = month['created_at'][:16].replace('T', ' ') if month else '2000-01-01 00:00'

        for mid in data[0].split():
            # Peek at headers — avoids marking as read
            _, hdr_data = mail.fetch(mid, '(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID SUBJECT DATE)])')
            hdr = email_lib.message_from_bytes(hdr_data[0][1])
            message_id = hdr.get('Message-ID', '').strip()
            subject = decode_str(hdr.get('Subject', ''))

            # Parse email date
            raw_date = hdr.get('Date', '')
            try:
                from email.utils import parsedate_to_datetime
                email_dt = parsedate_to_datetime(raw_date)
                # Convert to local system time for comparison with month_loaded_at (also local)
                email_dt_str = email_dt.astimezone().strftime('%Y-%m-%d %H:%M')
            except Exception:
                email_dt_str = '2099-01-01 00:00'

            # Skip emails that arrived before the month was loaded
            if email_dt_str < month_loaded_at:
                continue

            # Skip already processed
            if message_id and conn.execute(
                'SELECT 1 FROM processed_emails WHERE message_id=?', (message_id,)
            ).fetchone():
                continue

            if cfg['subject_filter'] and cfg['subject_filter'] not in subject:
                continue

            # Fetch full email without marking read
            _, full_data = mail.fetch(mid, '(BODY.PEEK[])')
            msg = email_lib.message_from_bytes(full_data[0][1])
            body = get_email_body(msg)
            # Skip the automated morning monitor submissions (uptime check on the website
            # forms). They carry a sentinel in the body, so this catches every form type —
            # even ones with no name/ID (e.g. 'מינוי סוכן').
            if any(m in body for m in ('MONITOR-CHECK-DO-NOT-PROCESS', 'automated-daily-check', 'monitor-check@example.com')):
                if message_id:
                    conn.execute(
                        'INSERT OR IGNORE INTO processed_emails (message_id, processed_at) VALUES (?,?)',
                        (message_id, datetime.datetime.now().isoformat())
                    )
                    conn.commit()
                continue
            fields = parse_renewal_email(body, subject)
            cid = process_renewal_data(fields, message_id=message_id,
                                        subject=subject, received_at=email_dt_str)
            # Mark processed regardless (matched or unmatched)
            if message_id:
                conn.execute(
                    'INSERT OR IGNORE INTO processed_emails (message_id, processed_at) VALUES (?,?)',
                    (message_id, datetime.datetime.now().isoformat())
                )
                conn.commit()
            if cid:
                _save_attachments(msg, cid)
            processed += 1

        conn.close()
        mail.logout()
    except Exception as e:
        print(f'[email-sync] שגיאה: {e}')

    return processed

POLICY_DOCS_DIR = os.path.join(ATTACHMENTS_DIR, 'policies')

def parse_harel_policy_pdf(source):
    """Best-effort field extraction from a Harel policy-schedule ('דף הרשימה') PDF page.
    `source` may be a file path or raw PDF bytes. Layout is consistent across doc types
    (new/renewal/cancellation/change) — same template, different coverage sections.
    Some fields (agent name, rare names with unusual glyphs) may need manual correction."""
    try:
        pdf_src = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
        with pdfplumber.open(pdf_src) as pdf:
            text = None
            for page in pdf.pages:
                t = page.extract_text() or ''
                if any('רשימה' in h for h in t.split('\n')[:2]):
                    text = t
                    break
        if not text:
            return {}
    except Exception as e:
        print(f'[policy-parse] שגיאת קריאת PDF: {e}')
        return {}

    lines = [get_display(l) for l in text.split('\n')]
    result = {}

    for i, l in enumerate(lines):
        m = re.search(r'\(([^0-9()]+)\s*(\d+)\)', l)
        if m and ('פוליסה' in l or 'תוספת' in l):
            result['doc_type_label'] = m.group(1).strip()
            result['doc_type_code'] = m.group(2)

        if i + 1 < len(lines) and ("מס' הפוליסה" in l or "מספר הפוליסה" in l):
            data_line = lines[i + 1]
            nums = re.findall(r'\d+', data_line.replace('/', ''))
            if nums:
                result['branch'] = nums[0]
            if len(nums) >= 3:
                result['agent_number'] = nums[2]
            agent_name = re.sub(r'[\d/\-]+', '', data_line).strip(' -()"\'')
            agent_name = re.sub(r'^[א-ת]\s+', '', agent_name)
            result['agent_name'] = agent_name.strip()

        if i + 1 < len(lines) and 'שם המבוטח' in l:
            result['insured_name'] = l.split('שם המבוטח וכתובתו')[-1].strip()
            addr_lines = []
            j = i + 1
            while j < len(lines) and 'תקופת' not in lines[j] and 'תאריך תחילת' not in lines[j]:
                addr_lines.append(lines[j].strip())
                j += 1
            result['address'] = ' '.join(addr_lines)

        if i + 1 < len(lines) and 'תקופת הביטוח' in l:
            m5 = re.findall(r'\d{2}/\d{2}/\d{4}', lines[i + 1])
            if len(m5) >= 2:
                result['period_start'] = m5[0]
                result['period_end'] = m5[1]

        if i + 1 < len(lines) and 'e-mail' in l:
            data_line = lines[i + 1]
            m6 = re.search(r'[\w.+-]+@[\w-]+\.[\w.]+', data_line)
            result['email'] = m6.group(0) if m6 else ''
            rest = data_line.replace(result['email'], '') if m6 else data_line
            phones = [p.replace(' ', '') for p in re.findall(r'0\d{1,2}-?\s?\d{6,7}', rest)]
            if phones:
                result['phone_mobile'] = phones[0]
            if len(phones) > 1:
                result['phone_home'] = phones[1]

        if i + 1 < len(lines) and 'ת.ז. מבוטח' in l:
            ids = re.findall(r'\d{7,9}', lines[i + 1])
            if ids:
                result['insured_id'] = ids[0]
            if len(ids) > 1:
                result['spouse_id'] = ids[1]

        if i + 1 < len(lines) and 'דמי ביטוח' in l and 'אשראי' in l:
            nums = re.findall(r'-?\d+\.\d{2}', lines[i + 1])
            if nums:
                result['premium'] = nums[0]
            if len(nums) > 1:
                result['total_payment'] = nums[-1]

    # Robust cancellation flag: the doc-type parenthetical can be mis-parsed, but a
    # cancellation reliably carries the "תוספת ביטול לפוליסה" header — trust that.
    full = '\n'.join(lines)
    if 'ביטול לפוליסה' in full or 'תוספת ביטול' in full:
        result['doc_type_label'] = 'ביטול'

    return result

_policy_check_lock = threading.Lock()

def check_policy_documents(days_back=30, keep_pdf=True):
    """Connect to IMAP, look for confirmed-policy emails (Harel ComposeDoc), extract the
    data, and (optionally) save the PDF. `days_back` widens the scan for backfills;
    `keep_pdf=False` parses in memory without storing the file (saves volume space)."""
    if not _policy_check_lock.acquire(blocking=False):
        print('[policy-docs] בדיקה כבר רצה — דילוג')
        return 0
    try:
        return _check_policy_documents_impl(days_back, keep_pdf)
    finally:
        _policy_check_lock.release()

def _check_policy_documents_impl(days_back=30, keep_pdf=True):
    cfg = EMAIL_CONFIG
    if not cfg['enabled'] or not cfg['imap_server'] or not cfg['password']:
        return 0

    from email.utils import parsedate_to_datetime
    processed = 0
    try:
        mail = imaplib.IMAP4_SSL(cfg['imap_server'], cfg['imap_port'])
        mail.login(cfg['username'], cfg['password'])
        mail.select('INBOX')

        since_date = (datetime.datetime.now() - datetime.timedelta(days=days_back)).strftime('%d-%b-%Y')
        status, data = mail.search(None, f'FROM "ComposeDoc@harel-ins.co.il" SINCE {since_date}')
        if status != 'OK':
            mail.logout()
            return 0

        conn = get_db()
        for mid in data[0].split():
            _, hdr_data = mail.fetch(mid, '(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID SUBJECT DATE)])')
            hdr = email_lib.message_from_bytes(hdr_data[0][1])
            message_id = hdr.get('Message-ID', '').strip()
            subject = decode_str(hdr.get('Subject', ''))
            try:
                doc_date = parsedate_to_datetime(hdr.get('Date', '')).astimezone().strftime('%Y-%m-%d %H:%M')
            except Exception:
                doc_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')

            if message_id and conn.execute(
                'SELECT 1 FROM policy_documents WHERE message_id=?', (message_id,)
            ).fetchone():
                continue

            m = re.search(r'(\d{6,})\s*$', subject.strip())
            policy_number = m.group(1) if m else None
            if not policy_number:
                continue

            customer = conn.execute(
                "SELECT id FROM customers WHERE ltrim(policy_number,'0')=?",
                (policy_number.lstrip('0'),)
            ).fetchone()
            customer_id = customer['id'] if customer else None

            _, full_data = mail.fetch(mid, '(BODY.PEEK[])')
            msg = email_lib.message_from_bytes(full_data[0][1])

            saved_any = False
            for part in msg.walk():
                cd = str(part.get('Content-Disposition', ''))
                if 'attachment' not in cd and part.get_content_type() != 'application/octet-stream':
                    continue
                raw_fn = part.get_filename()
                if not raw_fn:
                    continue
                filename = decode_str(raw_fn)
                data_bytes = part.get_payload(decode=True)
                if not data_bytes:
                    continue
                filepath = ''
                if keep_pdf:
                    folder_key = str(customer_id) if customer_id else f'unmatched_{policy_number}'
                    doc_dir = os.path.join(POLICY_DOCS_DIR, folder_key)
                    os.makedirs(doc_dir, exist_ok=True)
                    safe_fn = re.sub(r'[\\/*?:"<>|]', '_', filename)
                    filepath = os.path.join(doc_dir, safe_fn)
                    with open(filepath, 'wb') as f:
                        f.write(data_bytes)
                cur = conn.execute(
                    '''INSERT OR IGNORE INTO policy_documents
                       (customer_id, policy_number, filename, filepath, received_at, message_id)
                       VALUES (?,?,?,?,?,?)''',
                    (customer_id, policy_number, filename, filepath, doc_date, message_id)
                )
                conn.commit()
                saved_any = True
                status_label = f'ללקוח {customer_id}' if customer_id else 'לא זוהה לקוח'
                print(f'[policy-docs] {"נשמר" if keep_pdf else "עובד"}: {filename} ({policy_number}) {status_label}')

                if cur.lastrowid:
                    fields = parse_harel_policy_pdf(filepath if keep_pdf else data_bytes)
                    if fields:
                        conn.execute(
                            '''INSERT INTO policy_records
                               (policy_document_id, customer_id, policy_number, doc_type_label,
                                doc_type_code, branch, agent_name, agent_number, insured_name,
                                insured_id, spouse_id, address, phone_mobile, phone_home, email,
                                period_start, period_end, premium, total_payment, doc_date, extracted_at)
                               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                            (cur.lastrowid, customer_id, policy_number,
                             fields.get('doc_type_label'), fields.get('doc_type_code'),
                             fields.get('branch'), fields.get('agent_name'), fields.get('agent_number'),
                             fields.get('insured_name'), fields.get('insured_id'), fields.get('spouse_id'),
                             fields.get('address'), fields.get('phone_mobile'), fields.get('phone_home'),
                             fields.get('email'), fields.get('period_start'), fields.get('period_end'),
                             fields.get('premium'), fields.get('total_payment'), doc_date,
                             datetime.datetime.now().isoformat())
                        )
                        conn.commit()

            if saved_any:
                processed += 1

        conn.close()
        mail.logout()
    except Exception as e:
        print(f'[policy-docs] שגיאה: {e}')

    return processed

def email_poll_thread():
    """Background thread: check inbox every N seconds."""
    while True:
        time.sleep(EMAIL_CONFIG['check_interval'])
        try:
            n = check_email_inbox()
            if n:
                print(f'[email-sync] עובדו {n} מיילים חדשים')
        except Exception as e:
            print(f'[email-sync] שגיאת thread: {e}')
        try:
            n2 = check_policy_documents()
            if n2:
                print(f'[policy-docs] עובדו {n2} פוליסות חדשות')
        except Exception as e:
            print(f'[policy-docs] שגיאת thread: {e}')

# ── Admin email trigger ──────────────────────────────────────

@app.route('/admin/check-email', methods=['POST'])
@login_required
@admin_required
def admin_check_email():
    if not EMAIL_CONFIG['enabled']:
        flash('סנכרון מייל לא מוגדר עדיין — יש להגדיר IMAP בקובץ app.py', 'warning')
        return redirect(url_for('admin'))
    # Run in background so the page doesn't timeout on large inboxes
    threading.Thread(target=check_email_inbox, daemon=True).start()
    flash('בדיקת מייל הופעלה ברקע — רענן את הדף בעוד 10 שניות', 'info')
    return redirect(url_for('admin'))


@app.route('/refresh', methods=['POST'])
@login_required
def refresh_data():
    """Manual 'refresh' — pull emails + policy PDFs on demand, for when the
    background poll isn't running. Runs in a background thread so the request
    returns immediately (a synchronous IMAP scan exceeds gunicorn's worker
    timeout and gets the worker killed → 500)."""
    if not EMAIL_CONFIG['enabled']:
        flash('סנכרון מייל לא מוגדר עדיין', 'warning')
        return redirect(url_for('index'))

    def _run():
        try:
            check_email_inbox()
            check_policy_documents()
            conn = get_db()
            rebuild_insureds(conn)
            recompute_insured_statuses(conn)
            conn.close()
        except Exception as e:
            print(f'[refresh] שגיאה: {e}')

    threading.Thread(target=_run, daemon=True).start()
    flash('רענון הופעל — הנתונים יתעדכנו תוך מספר שניות. רענן את הדף.', 'info')
    return redirect(url_for('index'))


_backfill_state = {'running': False, 'done': 0, 'started': None, 'days': 0}

@app.route('/admin/backfill', methods=['POST'])
@login_required
@admin_required
def admin_backfill():
    """One-time backfill: scan up to a year of Harel PDFs, extract customer data +
    cancellations into the master. Data-only (keep_pdf=False) to stay within storage.
    Runs in the background; safe to leave and check back."""
    if _backfill_state['running']:
        flash('סריקה כבר רצה ברקע — המתן לסיומה', 'warning')
        return redirect(url_for('policy_records'))
    try:
        days = int(request.form.get('days', '30'))
    except ValueError:
        days = 30
    days = max(1, min(days, 400))

    def _run(days_back):
        _backfill_state.update(running=True, done=0, started=datetime.datetime.now().strftime('%H:%M'), days=days_back)
        try:
            n = check_policy_documents(days_back=days_back, keep_pdf=False)
            conn = get_db()
            rebuild_insureds(conn)
            recompute_insured_statuses(conn)
            conn.close()
            _backfill_state['done'] = n
            print(f'[backfill] הסתיים — {n} מסמכים חדשים, {days_back} ימים אחורה')
        except Exception as e:
            print(f'[backfill] שגיאה: {e}')
        finally:
            _backfill_state['running'] = False

    threading.Thread(target=_run, args=(days,), daemon=True).start()
    flash(f'סריקה אחורה של {days} ימים הופעלה ברקע — זה עשוי לקחת זמן. רענן את הדף מדי פעם.', 'info')
    return redirect(url_for('policy_records'))


@app.route('/submit', methods=['POST'])
def form_submit():
    """Direct POST from website forms (gaia-website / winner-website)."""
    data = request.get_json(silent=True) or {}
    if not data:
        return jsonify({'ok': False, 'error': 'no data'}), 400

    fields = {
        'name':           str(data.get('name') or '').strip(),
        'id_number':      normalize_id_number(data.get('id_number')),
        'phone':          str(data.get('phone') or '').strip(),
        'email':          str(data.get('email') or '').strip(),
        'installments':   str(data.get('installments') or '').strip(),
        'payment_method': str(data.get('payment_method') or '').strip(),
        'comments':       str(data.get('comments') or '').strip(),
        'brand':          str(data.get('brand') or '').strip(),
        'card_number':    str(data.get('card_number') or '').strip(),
        'card_expiry':    str(data.get('card_expiry') or '').strip(),
        'card_holder_id': str(data.get('card_holder_id') or '').strip(),
        'coverage_option': str(data.get('coverage_option') or '').strip(),
    }

    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    # Unique ID based on content + time to prevent exact duplicate submissions
    import hashlib
    unique_id = f"web-{hashlib.md5((fields['id_number']+fields['phone']+now).encode()).hexdigest()[:12]}"

    cid = process_renewal_data(fields, message_id=unique_id, subject=f"טופס חידוש {fields['brand']}", received_at=now)
    print(f'[submit] {fields["name"]} ({fields["id_number"]}) brand={fields["brand"]} → cid={cid}')
    return jsonify({'ok': True})


@app.route('/db-status')
def db_status():
    """Diagnostic endpoint — shows DB health"""
    try:
        conn = get_db()
        tables = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
        user_count = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0] if 'users' in tables else 'N/A'
        conn.close()
        return jsonify({
            'db_path': DB_PATH,
            'db_exists': os.path.exists(DB_PATH),
            'tables': tables,
            'user_count': user_count,
        })
    except Exception as e:
        return jsonify({'error': str(e), 'db_path': DB_PATH}), 500


# הפעל DB ו-email thread גם תחת gunicorn
try:
    print(f'[startup] calling init_db() on {DB_PATH}')
    init_db()
    print(f'[startup] init_db() done — db file exists: {os.path.exists(DB_PATH)}')
except Exception as _e:
    print(f'[startup] ERROR in init_db(): {_e}')
    import traceback; traceback.print_exc()

if EMAIL_CONFIG['enabled']:
    try:
        _t = threading.Thread(target=email_poll_thread, daemon=True)
        _t.start()
        print('[email-sync] Thread פעיל — יבדוק כל 5 דקות')
    except Exception as _e:
        print(f'[email-sync] ERROR starting thread: {_e}')

if __name__ == '__main__':
    print("=" * 50)
    print("מערכת שירות לקוחות פועלת!")
    print("כתובת גישה: http://localhost:5000")
    print("=" * 50)
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
