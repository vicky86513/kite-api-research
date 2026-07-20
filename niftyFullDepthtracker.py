"""
Live NIFTY current-month futures — FULL 5-level market depth tracker,
with cumulative volume tracking to spot where big size is resting.

Every tick sends 5 bid levels + 5 ask levels. Instead of just showing the
latest snapshot, this script ADDS the quantity seen at each exact price
onto a running total every time that price shows up in the depth — across
ALL 5 levels on both sides, not just the 5th level like the other futures
trackers. Over a few minutes this makes "walls" visible: a price where
huge cumulative quantity keeps showing up is a price big players are
consistently resting size at, which is what you're trying to spot.

Each price row is flagged as a BIG MONEY level automatically — no need to
hardcode a quantity threshold. It's flagged when its cumulative quantity
is at least BIG_ORDER_MULTIPLIER times the average cumulative quantity
across all tracked price levels for that side (bid average for bid rows,
ask average for ask rows). This adapts on its own to whatever lot sizes
are normal for NIFTY that session, instead of a fixed number that would
need retuning.

The dashboard also calls out the single biggest bid wall and biggest ask
wall directly in the stats strip — the two prices where the most
cumulative size has shown up on each side.

Automatically resolves the near-month NIFTY futures contract on NFO (no
need to hardcode the expiry each month), same as niftyCurrentMonthTracker.py.

Note: like the other trackers, this accumulates for the current day only —
restarting the script starts today's book fresh again (no rehydration).

Writes (date-stamped per day, matching the other trackers' convention):
    Excel/nifty_full_depth_<date>.xlsx / .json / .html

All log lines go into Logs/trading_logs_<date>.txt (shared with every
other tracker) via logger.py, tagged with [NIFTY-DEPTH].

Run:
    python niftyFullDepthTracker.py

Requires auth.py, instruments.py, and logger.py in the same folder.
"""

import os
import json as _json
import time
import threading
from datetime import datetime

from openpyxl import Workbook
from kiteconnect import KiteTicker

from auth import get_kite_session, API_KEY
from instruments import get_nearest_future
from logger import get_logger

log = get_logger("NIFTY-DEPTH")

EXCHANGE = "NFO"
UNDERLYING_NAME = "NIFTY"
OUTPUT_DIR = "Excel"
BASE_NAME = "nifty_full_depth"
SAVE_EVERY_SECONDS = 5
DEPTH_LEVELS = 5

# A price level is flagged "Big Money" when its cumulative bid (or ask)
# quantity is at least this many times the average cumulative quantity
# across all tracked price levels on that side. Raise it to flag only
# more extreme walls, lower it to flag more levels.
BIG_ORDER_MULTIPLIER = 3


class CumulativeDepthBook:
    """price -> {"bid_qty": cumulative int, "ask_qty": cumulative int,
    "last_updated": str}. Every appearance of a price anywhere in the 5
    bid/5 ask levels on a tick ADDS onto that price's running total —
    nothing is overwritten, so size that keeps reappearing at a price
    keeps stacking up and becomes visible as a wall."""

    @staticmethod
    def _dated_filepath(base_name, ext):
        date_str = datetime.now().strftime("%Y-%m-%d")
        return os.path.join(OUTPUT_DIR, f"{base_name}_{date_str}.{ext}")

    def __init__(self, base_name, title):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        self._base_name = base_name
        self.title = title
        self.lock = threading.Lock()
        self.current_date = datetime.now().date()
        self._set_filepaths()

        self.symbol = None
        self.ltp = None
        self.book = {}  # price -> {"bid_qty": int, "ask_qty": int, "last_updated": str}

    def _set_filepaths(self):
        self.filepath = self._dated_filepath(self._base_name, "xlsx")
        self.html_filepath = self._dated_filepath(self._base_name, "html")
        self.json_filepath = self._dated_filepath(self._base_name, "json")

    def _check_daily_rollover(self):
        today = datetime.now().date()
        if today != self.current_date:
            self.current_date = today
            self._set_filepaths()
            self.book = {}
            log.info(f"New day detected — starting fresh depth book: {self.filepath}")

    def set_symbol(self, symbol):
        self.symbol = symbol

    def update(self, ltp, buy_levels, sell_levels):
        """Adds the quantity at every bid level and every ask level from
        this tick onto each price's running cumulative total."""
        self._check_daily_rollover()
        with self.lock:
            self.ltp = ltp
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            for lvl in buy_levels[:DEPTH_LEVELS]:
                price = lvl.get("price")
                qty = lvl.get("quantity") or 0
                if price is None or qty == 0:
                    continue
                entry = self.book.setdefault(price, {"bid_qty": 0, "ask_qty": 0, "last_updated": None})
                entry["bid_qty"] += qty
                entry["last_updated"] = now

            for lvl in sell_levels[:DEPTH_LEVELS]:
                price = lvl.get("price")
                qty = lvl.get("quantity") or 0
                if price is None or qty == 0:
                    continue
                entry = self.book.setdefault(price, {"bid_qty": 0, "ask_qty": 0, "last_updated": None})
                entry["ask_qty"] += qty
                entry["last_updated"] = now

    def _big_money_thresholds(self):
        """Returns (bid_threshold, ask_threshold) — average cumulative qty
        per side times BIG_ORDER_MULTIPLIER. None if there's no data yet
        on that side (so nothing gets flagged prematurely)."""
        bid_values = [e["bid_qty"] for e in self.book.values() if e["bid_qty"] > 0]
        ask_values = [e["ask_qty"] for e in self.book.values() if e["ask_qty"] > 0]
        bid_threshold = (sum(bid_values) / len(bid_values)) * BIG_ORDER_MULTIPLIER if bid_values else None
        ask_threshold = (sum(ask_values) / len(ask_values)) * BIG_ORDER_MULTIPLIER if ask_values else None
        return bid_threshold, ask_threshold

    def flush(self):
        self._check_daily_rollover()
        with self.lock:
            self._write_xlsx()
            self._write_json()
            self._write_html()

    def _write_xlsx(self):
        bid_threshold, ask_threshold = self._big_money_thresholds()
        wb = Workbook()
        ws = wb.active
        ws.title = "CumulativeDepth"
        ws.append(["Price", "Bid Qty (cum.)", "Ask Qty (cum.)", "Net (Bid-Ask)",
                   "Total Qty", "Big Money Side", "Last Updated"])
        for price in sorted(self.book.keys()):
            entry = self.book[price]
            bid_qty, ask_qty = entry["bid_qty"], entry["ask_qty"]
            net = bid_qty - ask_qty
            total = bid_qty + ask_qty
            flags = []
            if bid_threshold and bid_qty >= bid_threshold:
                flags.append("BID")
            if ask_threshold and ask_qty >= ask_threshold:
                flags.append("ASK")
            ws.append([price, bid_qty, ask_qty, net, total, "/".join(flags), entry["last_updated"]])
        wb.save(self.filepath)

    def _write_json(self):
        bid_threshold, ask_threshold = self._big_money_thresholds()
        rows = []
        for price in sorted(self.book.keys()):
            entry = self.book[price]
            bid_qty, ask_qty = entry["bid_qty"], entry["ask_qty"]
            rows.append({
                "price": price,
                "bid_qty": bid_qty,
                "ask_qty": ask_qty,
                "net": bid_qty - ask_qty,
                "total_qty": bid_qty + ask_qty,
                "big_bid": bool(bid_threshold and bid_qty >= bid_threshold),
                "big_ask": bool(ask_threshold and ask_qty >= ask_threshold),
                "last_updated": entry["last_updated"],
            })
        payload = {
            "title": self.title,
            "symbol": self.symbol,
            "ltp": self.ltp,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "big_order_multiplier": BIG_ORDER_MULTIPLIER,
            "rows": rows,
        }
        with open(self.json_filepath, "w", encoding="utf-8") as f:
            _json.dump(payload, f, indent=2)

    def _write_html(self):
        now = datetime.now()
        bid_threshold, ask_threshold = self._big_money_thresholds()

        prices = sorted(self.book.keys())
        total_bid = sum(e["bid_qty"] for e in self.book.values())
        total_ask = sum(e["ask_qty"] for e in self.book.values())
        total_net = total_bid - total_ask
        net_class = "net-positive" if total_net > 0 else "net-negative" if total_net < 0 else "net-neutral"
        net_sign = "+" if total_net > 0 else ""

        biggest_bid_wall = max(self.book.items(), key=lambda kv: kv[1]["bid_qty"], default=(None, {"bid_qty": 0}))
        biggest_ask_wall = max(self.book.items(), key=lambda kv: kv[1]["ask_qty"], default=(None, {"ask_qty": 0}))
        big_bid_price, big_bid_entry = biggest_bid_wall
        big_ask_price, big_ask_entry = biggest_ask_wall

        rows_html = []
        for price in prices:
            entry = self.book[price]
            bid_qty, ask_qty = entry["bid_qty"], entry["ask_qty"]
            net = bid_qty - ask_qty
            row_net_class = "net-positive" if net > 0 else "net-negative" if net < 0 else "net-neutral"
            row_net_sign = "+" if net > 0 else ""

            is_big_bid = bool(bid_threshold and bid_qty >= bid_threshold)
            is_big_ask = bool(ask_threshold and ask_qty >= ask_threshold)

            fresh = ""
            if entry["last_updated"]:
                try:
                    updated_dt = datetime.strptime(entry["last_updated"], "%Y-%m-%d %H:%M:%S")
                    if (now - updated_dt).total_seconds() <= 10:
                        fresh = " fresh"
                except ValueError:
                    pass

            bid_badge = ' <span class="whale">&#128040; BIG</span>' if is_big_bid else ""
            ask_badge = ' <span class="whale">&#128040; BIG</span>' if is_big_ask else ""

            rows_html.append(f"""
                <tr class="row{fresh}{' big-row' if (is_big_bid or is_big_ask) else ''}">
                    <td class="price">{price}</td>
                    <td class="bid{' big-cell' if is_big_bid else ''}">{bid_qty:,}{bid_badge}</td>
                    <td class="ask{' big-cell' if is_big_ask else ''}">{ask_qty:,}{ask_badge}</td>
                    <td class="{row_net_class}">{row_net_sign}{net:,}</td>
                    <td class="time">{entry['last_updated'] or '-'}</td>
                </tr>""")
        rows_joined = "".join(rows_html) if rows_html else '<tr><td class="empty" colspan="5">Waiting for ticks...</td></tr>'

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="5">
<title>{self.title} &middot; Cumulative Depth</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
    :root {{
        --bg: #08090c; --bg-glow: #10141d; --panel: #11151d; --panel-2: #151a24;
        --border: #232a38; --text: #e8ecf1; --muted: #6f7a8a;
        --buy: #00d38a; --buy-dim: rgba(0,211,138,0.12);
        --sell: #ff4d6a; --sell-dim: rgba(255,77,106,0.12);
        --amber: #f5b53f; --amber-dim: rgba(245,181,63,0.12);
    }}
    * {{ box-sizing: border-box; }}
    body {{
        margin: 0; font-family: 'IBM Plex Mono', ui-monospace, monospace;
        background: radial-gradient(1100px 500px at 12% -10%, var(--bg-glow), transparent 60%), var(--bg);
        color: var(--text); padding: 44px 24px 60px;
    }}
    .wrap {{ max-width: 980px; margin: 0 auto; }}
    header {{ margin-bottom: 26px; }}
    .eyebrow {{ display: flex; align-items: center; gap: 8px; color: var(--amber);
        font-size: 11.5px; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 8px; }}
    .live-dot {{ width: 7px; height: 7px; border-radius: 50%; background: var(--amber); animation: pulse 1.8s infinite; }}
    @keyframes pulse {{
        0% {{ box-shadow: 0 0 0 0 rgba(245,181,63,0.55); }}
        70% {{ box-shadow: 0 0 0 8px rgba(245,181,63,0); }}
        100% {{ box-shadow: 0 0 0 0 rgba(245,181,63,0); }}
    }}
    h1 {{ font-family: 'Space Grotesk', sans-serif; font-size: 26px; margin: 0 0 4px; }}
    .subtitle {{ color: var(--muted); font-size: 12.5px; }}

    .stats {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px; margin-bottom: 14px; }}
    .stat {{ background: var(--panel); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; }}
    .stat .label {{ font-family: 'Space Grotesk', sans-serif; color: var(--muted); font-size: 10.5px;
        text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }}
    .stat .value {{ font-size: 18px; font-weight: 600; font-variant-numeric: tabular-nums; }}
    .stat .value.buy {{ color: var(--buy); }}
    .stat .value.sell {{ color: var(--sell); }}
    .net-positive {{ color: var(--buy); }}
    .net-negative {{ color: var(--sell); }}
    .net-neutral {{ color: var(--muted); }}

    .walls {{ display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 20px; }}
    .wall-card {{ border-radius: 10px; padding: 14px 16px; border: 1px solid var(--border); }}
    .wall-card.bid {{ background: var(--buy-dim); border-color: rgba(0,211,138,0.3); }}
    .wall-card.ask {{ background: var(--sell-dim); border-color: rgba(255,77,106,0.3); }}
    .wall-card .label {{ font-family: 'Space Grotesk', sans-serif; font-size: 10.5px; text-transform: uppercase;
        letter-spacing: 1px; margin-bottom: 6px; }}
    .wall-card.bid .label {{ color: var(--buy); }}
    .wall-card.ask .label {{ color: var(--sell); }}
    .wall-card .price {{ font-size: 22px; font-weight: 700; font-family: 'Space Grotesk', sans-serif; }}
    .wall-card .qty {{ font-size: 12px; color: var(--muted); margin-top: 2px; }}

    .panel {{ background: var(--panel); border: 1px solid var(--border); border-radius: 14px;
        overflow: hidden; box-shadow: 0 20px 50px rgba(0,0,0,0.4); }}
    table {{ width: 100%; border-collapse: collapse; }}
    thead th {{ text-align: right; font-family: 'Space Grotesk', sans-serif; font-size: 11px;
        text-transform: uppercase; letter-spacing: 0.8px; color: var(--muted);
        padding: 14px 18px; border-bottom: 1px solid var(--border); background: var(--panel-2); }}
    thead th:first-child {{ text-align: left; }}
    tbody td {{ text-align: right; padding: 12px 18px; border-bottom: 1px solid rgba(255,255,255,0.035);
        font-variant-numeric: tabular-nums; font-size: 13.5px; }}
    tbody tr:hover {{ background: rgba(255,255,255,0.02); }}
    tbody tr.fresh {{ background: rgba(245,181,63,0.05); }}
    tbody tr.big-row {{ background: var(--amber-dim); }}
    td.price {{ text-align: left; font-weight: 600; padding-left: 26px; }}
    td.bid {{ color: var(--buy); }}
    td.ask {{ color: var(--sell); }}
    td.bid.big-cell, td.ask.big-cell {{ font-weight: 700; }}
    .whale {{ display: inline-block; font-size: 10px; color: var(--amber); border: 1px solid var(--amber);
        border-radius: 4px; padding: 1px 5px; margin-left: 6px; vertical-align: middle; }}
    td.time {{ color: var(--muted); font-size: 11.5px; }}
    td.empty {{ text-align: center; color: var(--muted); padding: 48px 0; font-size: 13px; }}

    footer {{ text-align: center; color: var(--muted); font-size: 11px; margin-top: 18px; }}
</style>
</head>
<body>
<div class="wrap">
    <header>
        <div class="eyebrow"><span class="live-dot"></span>Live &middot; Cumulative full-depth (all 5 levels)</div>
        <h1>{self.title}</h1>
        <div class="subtitle">{self.symbol or ''} &middot; LTP {self.ltp if self.ltp is not None else '-'} &middot; Auto-refreshes every 5s &middot; {len(prices)} price levels tracked</div>
    </header>

    <div class="stats">
        <div class="stat"><div class="label">Total Bid Qty</div><div class="value buy">{total_bid:,}</div></div>
        <div class="stat"><div class="label">Total Ask Qty</div><div class="value sell">{total_ask:,}</div></div>
        <div class="stat"><div class="label">Net Imbalance</div><div class="value {net_class}">{net_sign}{total_net:,}</div></div>
        <div class="stat"><div class="label">Last Refresh</div><div class="value">{now.strftime('%H:%M:%S')}</div></div>
    </div>

    <div class="walls">
        <div class="wall-card bid">
            <div class="label">&#128040; Biggest Bid Wall</div>
            <div class="price">{big_bid_price if big_bid_price is not None else '-'}</div>
            <div class="qty">{big_bid_entry.get('bid_qty', 0):,} cumulative qty resting here</div>
        </div>
        <div class="wall-card ask">
            <div class="label">&#128040; Biggest Ask Wall</div>
            <div class="price">{big_ask_price if big_ask_price is not None else '-'}</div>
            <div class="qty">{big_ask_entry.get('ask_qty', 0):,} cumulative qty resting here</div>
        </div>
    </div>

    <div class="panel">
        <table>
            <thead>
                <tr>
                    <th>Price</th>
                    <th>Bid Qty (cum.)</th>
                    <th>Ask Qty (cum.)</th>
                    <th>Net (Bid &minus; Ask)</th>
                    <th>Last Updated</th>
                </tr>
            </thead>
            <tbody>{rows_joined}
            </tbody>
        </table>
    </div>
    <footer>&#128040; BIG = cumulative qty at least {BIG_ORDER_MULTIPLIER}x that side's average &middot; Generated by CumulativeDepthBook &middot; {now.strftime('%Y-%m-%d %H:%M:%S')}</footer>
</div>
</body>
</html>"""
        with open(self.html_filepath, "w", encoding="utf-8") as f:
            f.write(html)


def run_tracker():
    kite, access_token = get_kite_session()
    symbol, instrument_token = get_nearest_future(kite, EXCHANGE, UNDERLYING_NAME)
    log.info(f"Tracking {symbol} | Instrument token: {instrument_token}")

    book = CumulativeDepthBook(BASE_NAME, title="NIFTY Current-Month Futures — Cumulative Depth")
    book.set_symbol(symbol)

    kws = KiteTicker(API_KEY, access_token)

    def on_ticks(ws, ticks):
        for tick in ticks:
            if tick.get("instrument_token") != instrument_token:
                continue

            ltp = tick.get("last_price")

            # Market depth is only present in FULL mode: 5 buy + 5 sell levels.
            depth = tick.get("depth", {})
            buy_levels = depth.get("buy", [])
            sell_levels = depth.get("sell", [])

            log.info(
                f"{symbol} LTP: {ltp} | "
                f"Bids: {[(lvl['price'], lvl['quantity']) for lvl in buy_levels]} | "
                f"Asks: {[(lvl['price'], lvl['quantity']) for lvl in sell_levels]}"
            )

            book.update(ltp, buy_levels, sell_levels)

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
            book.flush()
            log.info("Cumulative depth Excel/JSON/HTML saved.")

    threading.Thread(target=periodic_save, daemon=True).start()

    log.info("Connecting to KiteTicker websocket...")
    kws.connect(threaded=False)  # blocking call — Ctrl+C to stop


if __name__ == "__main__":
    try:
        run_tracker()
    except KeyboardInterrupt:
        log.info("Stopped by user.")