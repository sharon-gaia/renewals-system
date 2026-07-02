import imaplib, email as email_lib, sys, datetime
from email.utils import parsedate_to_datetime
sys.stdout.reconfigure(encoding='utf-8')

mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
mail.login('sharon@gaia-ins.co.il', 'mieewohcyjygjfbx')
mail.select('INBOX')

status, data = mail.search(None, 'FROM "onboarding@resend.dev"')
ids = data[0].split()
print(f'Total emails from resend: {len(ids)}')

month_loaded_at = '2026-07-01 08:44'  # from DB

count_pass = 0
for mid in ids:
    _, hdr_data = mail.fetch(mid, '(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID SUBJECT DATE)])')
    hdr = email_lib.message_from_bytes(hdr_data[0][1])
    raw_date = hdr.get('Date', '')
    try:
        email_dt = parsedate_to_datetime(raw_date)
        email_dt_str = email_dt.astimezone().strftime('%Y-%m-%d %H:%M')
    except Exception as e:
        email_dt_str = '2099-01-01 00:00'
        print(f'  Date parse error: {e}')

    if email_dt_str >= month_loaded_at:
        count_pass += 1
        subj = hdr.get('Subject', '')
        print(f'  PASS: {email_dt_str} | {raw_date} | {subj[:40]}')

print(f'\nEmails passing date filter: {count_pass}')
mail.logout()
