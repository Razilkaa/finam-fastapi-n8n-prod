"""In-memory data storage for quotes_all."""

from __future__ import annotations

import time
from typing import Optional, TypedDict


class QuotesAllStore(TypedDict):
    quotes: list[dict]
    report_date: Optional[str]
    last_received_utc: Optional[str]


quotes_all_store: QuotesAllStore = {
    "quotes": [],
    "report_date": None,
    "last_received_utc": None,
}


def set_quotes_all(*, quotes: list[dict], report_date: Optional[str]) -> None:
    quotes_all_store["quotes"] = quotes
    quotes_all_store["report_date"] = report_date
    quotes_all_store["last_received_utc"] = time.strftime(
        "%Y-%m-%dT%H:%M:%SZ", time.gmtime()
    )
