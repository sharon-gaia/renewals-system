import sqlite3, sys
sys.stdout.reconfigure(encoding='utf-8')
conn = sqlite3.connect('renewals.db')
old = 'תגובה או וי כחול'
new = 'לקוח ענה/ V כחול'
conn.execute("UPDATE customers SET status=? WHERE status=?", (new, old))
conn.commit()
print('Done')
conn.close()
