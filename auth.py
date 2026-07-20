"""
Shared Kite Connect authentication helper.

Handles login and caches the access token for the day so you don't have to
paste the request token every time you run a tracker script.

Used by: mcx_crudeoil_tracker.py, nifty_current_month_tracker.py

.env file (same folder) should contain:
    KITE_API_KEY=your_api_key
    KITE_API_SECRET=your_api_secret
"""

import os
import json
from datetime import date

from dotenv import load_dotenv
from kiteconnect import KiteConnect

from logger import get_logger

load_dotenv()
log = get_logger("AUTH")

API_KEY = os.getenv("KITE_API_KEY")
API_SECRET = os.getenv("KITE_API_SECRET")

TOKEN_CACHE_FILE = "kite_session.json"


def load_cached_token():
    if not os.path.exists(TOKEN_CACHE_FILE):
        return None
    with open(TOKEN_CACHE_FILE) as f:
        cache = json.load(f)
    if cache.get("date") != str(date.today()):
        return None
    return cache.get("access_token")


def save_token_cache(access_token):
    with open(TOKEN_CACHE_FILE, "w") as f:
        json.dump({"date": str(date.today()), "access_token": access_token}, f)


def get_kite_session():
    """
    Returns (kite, access_token).
    Reuses today's cached token if it's still valid, otherwise walks through
    the login flow once and caches the new token for the rest of the day.
    """
    kite = KiteConnect(api_key=API_KEY)

    cached_token = load_cached_token()
    if cached_token:
        kite.set_access_token(cached_token)
        try:
            kite.profile()  # cheap call just to validate the cached token
            log.info("Using cached access token (no login needed today).")
            return kite, cached_token
        except Exception:
            log.warning("Cached token expired/invalid, need to log in again.")

    log.info(f"login_url: {kite.login_url()}")
    request_token = input("Enter request token: ").strip()
    data = kite.generate_session(request_token, api_secret=API_SECRET)
    access_token = data["access_token"]
    kite.set_access_token(access_token)
    save_token_cache(access_token)
    log.info("Authenticated successfully.")
    return kite, access_token