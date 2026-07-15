"""
Daily health-check report for the renewals system.
Sends a summary email to the admin every morning.
Run via Windows Task Scheduler at 08:00.
"""
import sqlite3
import smtplib
import datetime
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

DB_PATH = os.path.join(os.path.dirname(__file__), 'renewals.db')
REPORT_TO   = 'd.sharon.d@gmail.com'
REPORT_FROM = 'sharon@gaia-ins.co.il'
APP_PASSWORD = 'mieewohcyjygjfbx'
SYSTEM_URL  = 'https://negligee-pager-tray.ngrok-free.dev'

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def build_report():
    conn = get_db()
    today = datetime.date.today().isoformat()
    yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()

    month = conn.execute("SELECT * FROM months WHERE is_active=1").fetchone()
    month_name = month['name'] if month else '(אין חודש פעיל)'

    # Unmatched pending in admin queue
    unmatched = conn.execute(
        "SELECT * FROM unmatched_submissions WHERE status='pending' ORDER BY received_at DESC"
    ).fetchall()

    # Customers with form received (needs agent action)
    forms_pending = conn.execute(
        "SELECT * FROM customers WHERE month_id=(SELECT id FROM months WHERE is_active=1) "
        "AND status='טופס התקבל' ORDER BY form_received_at DESC"
    ).fetchall()

    # Customers marked 'דורש בירור' (needs admin)
    needs_clarify = conn.execute(
        "SELECT * FROM customers WHERE month_id=(SELECT id FROM months WHERE is_active=1) "
        "AND status='דורש בירור'"
    ).fetchall()

    # Overall stats
    stats = conn.execute(
        "SELECT COUNT(*) as total, "
        "SUM(CASE WHEN status='חודש' THEN 1 ELSE 0 END) as renewed, "
        "SUM(CASE WHEN status='' OR status IS NULL THEN 1 ELSE 0 END) as pending "
        "FROM customers WHERE month_id=(SELECT id FROM months WHERE is_active=1)"
    ).fetchone()

    conn.close()

    lines = []
    lines.append(f"דוח יומי — מערכת שירות לקוחות | {today}")
    lines.append(f"חודש פעיל: {month_name}")
    lines.append("")

    # Alert if anything needs attention
    alerts = []
    if unmatched:
        alerts.append(f"⚠️  {len(unmatched)} טפסים בתור אדמין שממתינים לבירור")
    if forms_pending:
        alerts.append(f"📋 {len(forms_pending)} לקוחות עם טופס שהתקבל — ממתין לטיפול נציג")
    if needs_clarify:
        alerts.append(f"❓ {len(needs_clarify)} לקוחות סומנו 'דורש בירור'")

    if alerts:
        lines.append("== דרוש טיפול ==")
        for a in alerts:
            lines.append(a)
        lines.append("")

    # Stats
    if stats:
        total = stats['total'] or 0
        renewed = stats['renewed'] or 0
        pending = stats['pending'] or 0
        pct = round(renewed / total * 100) if total else 0
        lines.append("== סטטיסטיקות חודש ==")
        lines.append(f"סה\"כ לקוחות: {total}")
        lines.append(f"חידשו: {renewed} ({pct}%)")
        lines.append(f"ממתינים לטיפול: {pending}")
        lines.append("")

    # Unmatched details
    if unmatched:
        lines.append("== תור אדמין — פרטים ==")
        for u in unmatched:
            lines.append(f"  • {u['name'] or '(ללא שם)'} | ת.ז: {u['id_number'] or '-'} | {u['received_at']}")
        lines.append("")

    # Forms pending details
    if forms_pending:
        lines.append("== טפסים שהתקבלו — ממתין לטיפול נציג ==")
        for c in forms_pending:
            lines.append(f"  • {c['name']} | {c['form_received_at']}")
        lines.append("")

    lines.append(f"כניסה למערכת: {SYSTEM_URL}")

    if not alerts:
        lines.insert(2, "✅ הכל תקין — אין פריטים ממתינים לטיפול")
        lines.insert(3, "")

    return "\n".join(lines)


def send_report(body):
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    today = datetime.date.today().strftime('%d/%m/%Y')
    has_alerts = '⚠️' in body or '📋' in body or '❓' in body
    subject = f"{'⚠️ ' if has_alerts else '✅ '}דוח מערכת שירות לקוחות — {today}"

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = REPORT_FROM
    msg['To'] = REPORT_TO
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(REPORT_FROM, APP_PASSWORD)
        s.sendmail(REPORT_FROM, REPORT_TO, msg.as_string())

    print(f"Report sent to {REPORT_TO}")


if __name__ == '__main__':
    body = build_report()
    print(body)
    print()
    send_report(body)
    print("Done.")
