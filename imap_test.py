import imaplib, email, sys
from email.utils import parsedate_to_datetime
sys.stdout.reconfigure(encoding='utf-8')

mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
mail.login('sharon@gaia-ins.co.il', 'mieewohcyjygjfbx')
mail.select('INBOX')

status, data = mail.search(None, 'FROM "onboarding@resend.dev"')
ids = data[0].split()
print(f'Found {len(ids)} emails from resend')

for mid in ids[-5:]:
    _, hdr = mail.fetch(mid, '(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID SUBJECT DATE)])')
    msg = email.message_from_bytes(hdr[0][1])
    raw_date = msg.get('Date', '')
    try:
        dt = parsedate_to_datetime(raw_date)
        dt_str = dt.strftime('%Y-%m-%d %H:%M')
    except:
        dt_str = 'parse error'
    print('---')
    print('Subject:', msg.get('Subject',''))
    print('Date raw:', raw_date)
    print('Date parsed:', dt_str)
    print('Msg-ID:', msg.get('Message-ID','')[:60])

mail.logout()
