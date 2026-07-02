"""
worker.py — סורק מיילים ברקע (Railway worker process)
מריץ check_email_inbox כל 5 דקות, 24/7
"""
import time
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

from app import check_email_inbox, init_db, EMAIL_CONFIG

if __name__ == '__main__':
    init_db()
    print('[worker] מתחיל — יסרוק מיילים כל 5 דקות')

    # סריקה ראשונה מיד בהפעלה (יתפוס את כל המיילים הממתינים)
    try:
        n = check_email_inbox()
        print(f'[worker] סריקה ראשונה — עובדו {n} מיילים')
    except Exception as e:
        print(f'[worker] שגיאה בסריקה ראשונה: {e}')

    while True:
        time.sleep(EMAIL_CONFIG['check_interval'])
        try:
            n = check_email_inbox()
            if n:
                print(f'[worker] עובדו {n} מיילים חדשים')
        except Exception as e:
            print(f'[worker] שגיאה: {e}')
