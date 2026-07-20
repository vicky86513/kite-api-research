"""
Live MCX Crude Oil futures — 5th-level bid/ask price-book tracker.

Automatically resolves the near-month CRUDEOIL futures contract on MCX (no
need to hardcode the expiry each month — it just picks whichever contract
has the nearest expiry >= today). Writes to crude_oil_price_book.xlsx
(+ matching .json/.html), one row per unique strike price, accumulating
bid/ask quantity and refreshing the timestamp.

All log lines from this script go into trading_logs.txt (shared with the
Nifty/Sensex trackers and auth.py) via logger.py, tagged with [MCX].

Run:
    python mcx_crudeoil_tracker.py

Requires auth.py, price_book.py, and logger.py in the same folder.
"""

import time
import threading

from kiteconnect import KiteTicker

from auth import get_kite_session, API_KEY
from price_book import PriceBookWriter
from instruments import get_nearest_future
from logger import get_logger

log = get_logger("MCX")

EXCHANGE = "MCX"
UNDERLYING_NAME = "CRUDEOIL"
EXCEL_FILE = "crude_oil_price_book.xlsx"
SAVE_EVERY_SECONDS = 5


def run_tracker():
    kite, access_token = get_kite_session()
    symbol, instrument_token = get_nearest_future(kite, EXCHANGE, UNDERLYING_NAME)
    log.info(f"Tracking {symbol} | Instrument token: {instrument_token}")

    writer = PriceBookWriter(EXCEL_FILE, title="MCX Crude Oil Futures")
    kws = KiteTicker(API_KEY, access_token)

    def on_ticks(ws, ticks):
        for tick in ticks:
            price = tick.get("last_price")

            # Market depth is only present in FULL mode: 5 buy + 5 sell levels.
            depth = tick.get("depth", {})
            buy_levels = depth.get("buy", [])
            sell_levels = depth.get("sell", [])

            bid5_price = buy_levels[4]["price"] if len(buy_levels) >= 5 else None
            bid5_qty = buy_levels[4]["quantity"] if len(buy_levels) >= 5 else None
            ask5_price = sell_levels[4]["price"] if len(sell_levels) >= 5 else None
            ask5_qty = sell_levels[4]["quantity"] if len(sell_levels) >= 5 else None

            log.info(
                f"{symbol} Price: {price} | "
                f"5th Bid: {bid5_price} x {bid5_qty} | "
                f"5th Ask: {ask5_price} x {ask5_qty}"
            )

            writer.update(bid5_price, bid_qty=bid5_qty)
            writer.update(ask5_price, ask_qty=ask5_qty)

    def on_connect(ws, response):
        log.info(f"WebSocket connected. Subscribing to {instrument_token}.")
        ws.subscribe([instrument_token])
        ws.set_mode(ws.MODE_FULL, [instrument_token])

    def on_close(ws, code, reason):
        log.warning(f"WebSocket closed: {code} {reason}")

    def on_error(ws, code, reason):
        log.error(f"WebSocket error: {code} {reason}")

    kws.on_ticks = on_ticks
    kws.on_connect = on_connect
    kws.on_close = on_close
    kws.on_error = on_error

    def periodic_save():
        while True:
            time.sleep(SAVE_EVERY_SECONDS)
            writer.flush()
            log.info("Excel/JSON/HTML price book saved.")

    threading.Thread(target=periodic_save, daemon=True).start()

    log.info("Connecting to KiteTicker websocket...")
    kws.connect(threaded=False)  # blocking call — Ctrl+C to stop


if __name__ == "__main__":
    try:
        run_tracker()
    except KeyboardInterrupt:
        log.info("Stopped by user.")