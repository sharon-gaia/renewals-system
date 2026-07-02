import sqlite3, shutil, os, sys
sys.stdout = __import__('io').TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

conn = sqlite3.connect('renewals.db')
conn.execute('DELETE FROM customers')
conn.execute('DELETE FROM months')
conn.execute('DELETE FROM processed_emails')
try:
    conn.execute('DELETE FROM customer_attachments')
except: pass
conn.execute('''
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
    )
''')
conn.commit()
conn.close()

if os.path.exists('attachments'):
    shutil.rmtree('attachments')

print('cleared OK')
