"""
Daily health-check report for the renewals system.
Fetches the summary from the PRODUCTION cloud app (Railway) and emails it to the admin.
(Previously read a LOCAL dev SQLite copy — which showed stale/dev numbers, not production.)
Run via Windows Task Scheduler at 08:00.
"""
import smtplib
import datetime
import sys
import os
import json
from urllib import request as urlrequest

sys.stdout.reconfigure(encoding='utf-8')

REPORT_TO    = 'd.sharon.d@gmail.com'
REPORT_FROM  = 'sharon@gaia-ins.co.il'
APP_PASSWORD = 'mieewohcyjygjfbx'
CLOUD_URL    = os.environ.get('RENEWALS_URL', 'https://renewals-system-production.up.railway.app').rstrip('/')


def _token():
    """Bot API token. From env, else an untracked local report_token.txt (gitignored)."""
    t = os.environ.get('WA_API_TOKEN')
    if t:
        return t.strip()
    path = os.path.join(os.path.dirname(__file__), 'report_token.txt')
    with open(path, encoding='utf-8') as f:
        return f.read().strip()


def fetch_report():
    """Pull the server-computed report text from the production app."""
    req = urlrequest.Request(CLOUD_URL + '/api/daily-report', headers={'X-WA-Token': _token()})
    with urlrequest.urlopen(req, timeout=30) as r:
        data = json.loads(r.read().decode('utf-8'))
    return data['report']


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
    try:
        body = fetch_report()
    except Exception as e:
        body = (f"⚠️ הדוח היומי נכשל בשליפת נתונים מהענן: {e}\n"
                f"בדוק את {CLOUD_URL}/api/daily-report ואת report_token.txt")
    print(body)
    print()
    send_report(body)
    print("Done.")
