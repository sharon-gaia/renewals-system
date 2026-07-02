import sqlite3, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = sqlite3.connect('renewals.db')
conn.row_factory = sqlite3.Row

print('=== customers ===')
for r in conn.execute("SELECT id, name, status, form_received_at FROM customers ORDER BY id DESC LIMIT 10").fetchall():
    print(dict(r))

print('\n=== unmatched_submissions ===')
for r in conn.execute("SELECT * FROM unmatched_submissions ORDER BY id DESC LIMIT 10").fetchall():
    print(dict(r))

print('\n=== processed_emails ===')
for r in conn.execute("SELECT * FROM processed_emails ORDER BY processed_at DESC LIMIT 5").fetchall():
    print(dict(r))
