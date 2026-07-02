"""
health_check.py — Flask Blueprint לבדיקת תקינות המערכת
======================================================
הוספה לאפליקציה הראשית:

    from health_check import health_bp
    app.register_blueprint(health_bp)

משתנה סביבה נדרש (נוסף על הקיימים):
    RESEND_API_KEY   — מפתח ה-API של Resend
"""

import os
import time
from datetime import datetime, timezone

import requests
from flask import Blueprint, jsonify

health_bp = Blueprint("health", __name__)

SITES = {
    "gaia": "https://www.gaia-ins.co.il",
    "winner": "https://www.winner-ins.co.il",
}

RESEND_API_URL = "https://api.resend.com/emails"


# ──────────────────────────────────────────────
# פונקציות עזר
# ──────────────────────────────────────────────

def _check_site(name: str, url: str) -> dict:
    """בדיקת זמינות אתר — HEAD request עם timeout של 8 שניות."""
    start = time.monotonic()
    try:
        resp = requests.head(url, timeout=8, allow_redirects=True)
        ms = round((time.monotonic() - start) * 1000)
        ok = resp.status_code < 400
        return {
            "status": "ok" if ok else "error",
            "http_status": resp.status_code,
            "response_time_ms": ms,
        }
    except requests.Timeout:
        return {"status": "error", "error": "timeout"}
    except requests.RequestException as e:
        return {"status": "error", "error": str(e)}


def _check_resend() -> dict:
    """בדיקת תקינות Resend API — GET /emails (לא שולח מייל)."""
    api_key = os.getenv("RESEND_API_KEY", "")
    if not api_key:
        return {"status": "skipped", "reason": "RESEND_API_KEY not set"}

    try:
        resp = requests.get(
            "https://api.resend.com/domains",
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=8,
        )
        if resp.status_code == 200:
            return {"status": "ok"}
        return {"status": "error", "http_status": resp.status_code}
    except requests.RequestException as e:
        return {"status": "error", "error": str(e)}


# ──────────────────────────────────────────────
# Endpoint
# ──────────────────────────────────────────────

@health_bp.route("/health/forms", methods=["GET"])
def health_forms():
    """
    GET /health/forms
    מחזיר סטטוס כולל + פירוט לכל רכיב.
    HTTP 200  → הכל תקין
    HTTP 503  → לפחות בדיקה אחת נכשלה
    """
    checks = {}

    # בדיקת אתרים
    for name, url in SITES.items():
        checks[f"{name}_site"] = _check_site(name, url)

    # בדיקת Resend
    checks["resend_api"] = _check_resend()

    # סטטוס כולל
    failed = [k for k, v in checks.items() if v.get("status") == "error"]
    overall = "ok" if not failed else "error"

    payload = {
        "status": overall,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "checks": checks,
    }
    if failed:
        payload["failed"] = failed

    return jsonify(payload), (200 if overall == "ok" else 503)


@health_bp.route("/health", methods=["GET"])
def health_simple():
    """בדיקה פשוטה שה-backend רץ — תמיד מחזיר 200."""
    return jsonify({"status": "ok", "timestamp": datetime.now(timezone.utc).isoformat()})
