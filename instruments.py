"""
Shared instrument-resolution helper.

Finds the near-month futures contract for a given underlying on a given
exchange (nearest expiry >= today), so no tracker script needs to hardcode
a contract symbol that goes stale every month/expiry.

Used by: mcx_crudeoil_tracker.py, nifty_current_month_tracker.py,
         sensex_current_month_tracker.py
"""

from datetime import date


def get_nearest_future(kite, exchange, name):
    """
    Returns (tradingsymbol_with_exchange, instrument_token) for the
    nearest-expiry FUT contract matching `name` on `exchange`.

    Examples:
        get_nearest_future(kite, "MCX", "CRUDEOIL")  -> MCX crude oil future
        get_nearest_future(kite, "NFO", "NIFTY")      -> NIFTY future
        get_nearest_future(kite, "BFO", "SENSEX")     -> SENSEX future
    """
    instruments = kite.instruments(exchange)
    futures = [
        i for i in instruments
        if i["name"] == name
        and i["instrument_type"] == "FUT"
        and i["expiry"] >= date.today()
    ]
    if not futures:
        raise RuntimeError(f"No {name} futures contracts found on {exchange}.")
    nearest = min(futures, key=lambda i: i["expiry"])
    symbol = f"{exchange}:{nearest['tradingsymbol']}"
    return symbol, nearest["instrument_token"]