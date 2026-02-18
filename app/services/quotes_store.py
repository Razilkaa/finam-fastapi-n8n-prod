"""In-memory data storage for quotes."""

from __future__ import annotations

import time
from typing import TypedDict, Optional


class QuotesStore(TypedDict):
    quotes: list[dict]
    report_date: Optional[str]
    last_received_utc: Optional[str]


quotes_store: QuotesStore = {
    "quotes": [],
    "report_date": None,
    "last_received_utc": None,
}


def set_quotes(*, quotes: list[dict], report_date: Optional[str]) -> None:
    quotes_store["quotes"] = quotes
    quotes_store["report_date"] = report_date
    quotes_store["last_received_utc"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())

