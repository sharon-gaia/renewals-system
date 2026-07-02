import sys, os, imaplib, email as email_lib, datetime, sqlite3
from email.utils import parsedate_to_datetime
sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

import app as flask_app

# Monkey-patch to add verbose logging
original = flask_app.check_email_inbox

def verbose_check():
    cfg = flask_app.EMAIL_CONFIG
    print(f'Config: enabled={cfg["enabled"]}, server={cfg["imap_server"]}')
    try:
        mail = imaplib.IMAP4_SSL(cfg['imap_server'], cfg['imap_port'])
        mail.login(cfg['username'], cfg['password'])
        mail.select('INBOX')
        status, data = mail.search(None, f'FROM "{cfg["sender_filter"]}"')
        print(f'Search status: {status}, count: {len(data[0].split())}')

        conn = sqlite3.connect('renewals.db')
        conn.row_factory = sqlite3.Row
        month = conn.execute("SELECT created_at FROM months WHERE is_active=1 ORDER BY id DESC LIMIT 1").fetchone()
        month_loaded_at = month['created_at'][:16] if month else '2000-01-01 00:00'
        print(f'Month loaded at: {month_loaded_at}')

        processed = 0
        for mid in data[0].split():
            _, hdr_data = mail.fetch(mid, '(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID SUBJECT DATE)])')
            hdr = email_lib.message_from_bytes(hdr_data[0][1])
            message_id = hdr.get('Message-ID', '').strip()
            raw_date = hdr.get('Date', '')
            try:
                email_dt = parsedate_to_datetime(raw_date)
                email_dt_str = email_dt.astimezone().strftime('%Y-%m-%d %H:%M')
            except Exception as e:
                email_dt_str = '2099-01-01 00:00'
                print(f'  Date parse error: {e}')

            if email_dt_str < month_loaded_at:
                continue

            already = conn.execute('SELECT 1 FROM processed_emails WHERE message_id=?', (message_id,)).fetchone()
            if already:
                print(f'  Already processed: {message_id[:30]}')
                continue

            print(f'  NEW EMAIL: {email_dt_str} | msgid: {message_id[:40]}')
            processed += 1

        conn.close()
        mail.logout()
        print(f'Total processed: {processed}')
        return processed
    except Exception as e:
        import traceback
        print(f'ERROR: {e}')
        traceback.print_exc()
        return 0

with flask_app.app.app_context():
    verbose_check()
