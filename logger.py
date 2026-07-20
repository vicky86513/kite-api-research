"""
Shared logger — every script (auth, MCX tracker, Nifty tracker, Sensex
tracker, options tracker) logs into ONE common file per day, so you can
open a single file and see everything that happened across the whole
system in chronological order, instead of hunting through separate log
files per script.

The log file is date-stamped, e.g. Logs/trading_logs_2026-07-17.txt, and
automatically rolls over to a fresh file the moment the calendar date
changes — even if a script happens to keep running past midnight, no
restart needed. The Logs/ folder is created automatically if it doesn't
exist.

Usage:
    from logger import get_logger
    log = get_logger("MCX")
    log.info("Something happened")
    log.error("Something went wrong")
"""

import os
import logging
from datetime import date

LOG_DIR = "Logs"

_configured_names = set()


def _dated_log_path():
    return os.path.join(LOG_DIR, f"trading_logs_{date.today()}.txt")


class DailyFileHandler(logging.Handler):
    """Minimal daily-rotating file handler. Writes to
    Logs/trading_logs_<today>.txt and automatically switches to a fresh,
    freshly-dated file the instant the calendar date changes."""

    def __init__(self, formatter):
        super().__init__()
        self.setFormatter(formatter)
        self._current_date = None
        self._stream = None
        self._open_for_today()

    def _open_for_today(self):
        os.makedirs(LOG_DIR, exist_ok=True)
        if self._stream:
            self._stream.close()
        self._current_date = date.today()
        self._stream = open(_dated_log_path(), "a", encoding="utf-8")

    def emit(self, record):
        if date.today() != self._current_date:
            self._open_for_today()
        try:
            self._stream.write(self.format(record) + "\n")
            self._stream.flush()
        except Exception:
            self.handleError(record)


def get_logger(name):
    """Returns a logger tagged with `name` (e.g. 'MCX', 'NIFTY', 'AUTH').
    All loggers write to the same day's Logs/trading_logs_<date>.txt file,
    plus the console."""
    logger = logging.getLogger(name)

    if name in _configured_names:
        return logger  # already wired up — avoid attaching duplicate handlers

    logger.setLevel(logging.INFO)
    logger.propagate = False

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-7s | %(name)-6s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = DailyFileHandler(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    _configured_names.add(name)
    return logger