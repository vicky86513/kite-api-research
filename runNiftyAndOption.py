"""
Convenience launcher — runs the NIFTY current-month futures tracker AND the
options tracker (CE/PE, from OPTIONS in options_tracker.py) at the same
time, in a single process, using threads.

Logs into Kite once up front so you're only prompted for a request token
once (both trackers below then just reuse the cached token), then starts
both trackers concurrently — each keeps its own separate KiteTicker
websocket connection and writes to its own files under Excel/ and Logs/,
same as if you'd run them in two separate terminals.

Run:
    python run_nifty_and_options.py

Stop with Ctrl+C (stops both trackers).

Requires auth.py, price_book.py, logger.py, niftyCurrentMonthTracker.py,
and optionstracker.py in the same folder.
"""

import threading

from auth import get_kite_session
from logger import get_logger
import niftyCurrentMonthTracker
import optionstracker

log = get_logger("LAUNCHER")


def main():
    # Authenticate once here so both trackers below just reuse today's
    # cached token instead of each separately prompting for a request token.
    get_kite_session()

    threads = [
        threading.Thread(target=niftyCurrentMonthTracker.run_tracker, name="NiftyTracker", daemon=True),
        threading.Thread(target=optionstracker.run_tracker, name="OptionsTracker", daemon=True),
    ]

    for t in threads:
        t.start()
        log.info(f"Started {t.name}")

    try:
        # Keep the main thread alive while both trackers run in the background.
        while any(t.is_alive() for t in threads):
            for t in threads:
                t.join(timeout=1)
    except KeyboardInterrupt:
        log.info("Ctrl+C received — stopping (both trackers are daemon threads and will exit with this process).")


if __name__ == "__main__":
    main()