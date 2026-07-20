"""
Live options — 5th-level bid/ask price-book tracker, generic across any
underlying (MCX CRUDEOIL, NFO NIFTY, BFO SENSEX, or anything else you add).

Add/remove/edit entries in the OPTIONS list below — each entry is one
strike, and this script automatically tracks BOTH its CE and PE legs.

IMPORTANT: unlike the futures trackers, CE/PE contracts are NOT specified
by a hardcoded tradingsymbol string. You just give the underlying + strike,
and this script looks up the actual nearest-expiry contract from Kite's
live instrument list for that exchange. This is what gets you the WEEKLY
expiry automatically — a hardcoded symbol like "NIFTY26JUL24200CE" always
means the monthly contract (3-letter month code = monthly convention),
which is why that approach was returning monthly data instead of weekly.
Picking the contract with the nearest expiry >= today naturally lands on
the next weekly expiry, since weeklies occur every week (the monthly
expiry is just that month's last weekly, so it's never nearer).

ALL contracts across ALL entries are subscribed on ONE single KiteTicker
websocket connection, but each CE/PE leg gets its own separate price book
so they show up as separate tabs on the dashboard instead of being mixed
together.

Writes (per entry, using its file_prefix), date-stamped per day:
    Excel/<file_prefix>_ce_price_book_<date>.xlsx / .json / .html
    Excel/<file_prefix>_pe_price_book_<date>.xlsx / .json / .html

All log lines go into Logs/trading_logs_<date>.txt tagged [OPTIONS].

Run:
    python options_tracker.py

Requires auth.py, price_book.py, and logger.py in the same folder.
"""

import os
import json
import time
import threading
from datetime import date, datetime

from kiteconnect import KiteTicker

from auth import get_kite_session, API_KEY
from price_book import PriceBookWriter, OUTPUT_DIR
from logger import get_logger

log = get_logger("OPTIONS")

# Each entry: one strike on one underlying/exchange. `file_prefix` controls
# the output filenames (Excel/<file_prefix>_ce_price_book_<date>.*  and
# _pe_...). `strike` must match exactly how Kite lists it (usually a plain
# number, e.g. 24200).
# >>> EDIT THIS LIST for your strikes — add as many as you want.
# Comment an entry out (or delete it) to stop tracking it. Whatever's active
# here is automatically written to Excel/options_manifest.json on startup,
# and dashboard.html reads that file to build its CE/PE tabs — so the
# dashboard always matches this list with no manual editing over there. <<<
OPTIONS = [
    # {
    #     "label": "CRUDEOIL 7650",
    #     "file_prefix": "crudeoil_7650",
    #     "exchange": "MCX",
    #     "underlying": "CRUDEOIL",
    #     "strike": 7650,
    # },
    {
        "label": "NIFTY 24200",
        "file_prefix": "nifty_24200",
        "exchange": "NFO",
        "underlying": "NIFTY",
        "strike": 24200,
    },
    {
        "label": "NIFTY 23800",
        "file_prefix": "nifty_23800",
        "exchange": "NFO",
        "underlying": "NIFTY",
        "strike": 23800,
    },

    # {
    #     "label": "SENSEX 81500",
    #     "file_prefix": "sensex_81500",
    #     "exchange": "BFO",
    #     "underlying": "SENSEX",
    #     "strike": 81500,
    # },
    # Add more strikes/underlyings here the same way — no other code changes needed.
]

SAVE_EVERY_SECONDS = 5
MANIFEST_FILE = os.path.join(OUTPUT_DIR, "options_manifest.json")


def write_options_manifest():
    """Writes Excel/options_manifest.json listing exactly the entries
    currently active in OPTIONS above, so dashboard.html can dynamically
    build one CE tab + one PE tab per entry."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    manifest = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "options": [
            {
                "label": opt["label"],
                "file_prefix": opt["file_prefix"],
                "ce_file": f"{opt['file_prefix']}_ce_price_book.json",
                "pe_file": f"{opt['file_prefix']}_pe_price_book.json",
            }
            for opt in OPTIONS
        ],
    }
    with open(MANIFEST_FILE, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    log.info(f"Wrote options manifest with {len(OPTIONS)} entries to {MANIFEST_FILE}")


def get_nearest_weekly_option(instruments, exchange, underlying, strike, option_type):
    """Finds the option contract (CE or PE) for `underlying`/`strike` on
    `exchange` with the nearest expiry >= today. Since weekly expiries
    happen every week (the monthly is just the last weekly of the month),
    "nearest" is always the next weekly expiry, not the monthly one."""
    candidates = [
        i for i in instruments
        if i["name"] == underlying
        and i["instrument_type"] == option_type
        and float(i["strike"]) == float(strike)
        and i["expiry"] >= date.today()
    ]
    if not candidates:
        raise RuntimeError(
            f"No {option_type} contract found for {underlying} strike {strike} on {exchange}. "
            f"Check that the strike matches an actual listed strike."
        )
    nearest = min(candidates, key=lambda i: i["expiry"])
    symbol = f"{exchange}:{nearest['tradingsymbol']}"
    return symbol, nearest["instrument_token"], nearest["expiry"]


def build_routing_table(kite):
    """Resolves every CE/PE leg above to its nearest-weekly-expiry
    instrument token by looking up Kite's live instrument list per
    exchange, and builds a PriceBookWriter for each leg.
    Returns {instrument_token: (option_type, symbol, writer)}."""
    instruments_by_exchange = {}
    for opt in OPTIONS:
        exch = opt["exchange"]
        if exch not in instruments_by_exchange:
            instruments_by_exchange[exch] = kite.instruments(exch)

    routing = {}
    for opt in OPTIONS:
        instruments = instruments_by_exchange[opt["exchange"]]

        ce_symbol, ce_token, ce_expiry = get_nearest_weekly_option(
            instruments, opt["exchange"], opt["underlying"], opt["strike"], "CE"
        )
        pe_symbol, pe_token, pe_expiry = get_nearest_weekly_option(
            instruments, opt["exchange"], opt["underlying"], opt["strike"], "PE"
        )

        ce_writer = PriceBookWriter(
            f"{opt['file_prefix']}_ce_price_book.xlsx",
            title=f"{opt['label']} CE ({ce_expiry.strftime('%d-%b')})",
            price_step=None,  # no rounding — exact option premium prices
        )
        pe_writer = PriceBookWriter(
            f"{opt['file_prefix']}_pe_price_book.xlsx",
            title=f"{opt['label']} PE ({pe_expiry.strftime('%d-%b')})",
            price_step=None,  # no rounding — exact option premium prices
        )

        routing[ce_token] = ("CE", ce_symbol, ce_writer)
        routing[pe_token] = ("PE", pe_symbol, pe_writer)

        log.info(f"Resolved {opt['label']}: CE={ce_symbol} (token {ce_token}, expiry {ce_expiry}), "
                 f"PE={pe_symbol} (token {pe_token}, expiry {pe_expiry})")

    return routing


def run_tracker():
    write_options_manifest()
    kite, access_token = get_kite_session()
    routing = build_routing_table(kite)
    all_tokens = list(routing.keys())

    kws = KiteTicker(API_KEY, access_token)

    def on_ticks(ws, ticks):
        for tick in ticks:
            entry = routing.get(tick.get("instrument_token"))
            if not entry:
                continue
            option_type, symbol, writer = entry

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
                f"[{option_type}] {symbol} Price: {price} | "
                f"5th Bid: {bid5_price} x {bid5_qty} | "
                f"5th Ask: {ask5_price} x {ask5_qty}"
            )

            writer.update(bid5_price, bid_qty=bid5_qty)
            writer.update(ask5_price, ask_qty=ask5_qty)

    def on_connect(ws, response):
        log.info(f"WebSocket connected. Subscribing to {len(all_tokens)} contracts.")
        ws.subscribe(all_tokens)
        ws.set_mode(ws.MODE_FULL, all_tokens)

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
            for _, _, writer in routing.values():
                writer.flush()
            log.info("All CE/PE Excel/JSON/HTML price books saved.")

    threading.Thread(target=periodic_save, daemon=True).start()

    log.info("Connecting to KiteTicker websocket...")
    kws.connect(threaded=False)  # blocking call — Ctrl+C to stop


if __name__ == "__main__":
    try:
        run_tracker()
    except KeyboardInterrupt:
        log.info("Stopped by user.")