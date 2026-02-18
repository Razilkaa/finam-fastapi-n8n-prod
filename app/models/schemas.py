"""Pydantic schemas for API requests and responses."""
from typing import Optional
from pydantic import BaseModel, ConfigDict


class CalendarEvent(BaseModel):
    """Schema for calendar event."""
    date: str
    time: Optional[str] = ""
    country: str
    event: Optional[str] = None
    holiday: Optional[str] = None
    Key: Optional[int | str] = 0
    source_id: Optional[str] = None


class EventsPayload(BaseModel):
    """Schema for events payload."""
    events: list[CalendarEvent]


class StatusResponse(BaseModel):
    """Schema for status response."""
    status: str
    data: dict[str, int]


class ReceiveResponse(BaseModel):
    """Schema for receive endpoint response."""
    status: str
    total_received: int
    split: dict[str, int]


class QuoteItem(BaseModel):
    """Schema for a single quote item (fields are intentionally flexible)."""
    model_config = ConfigDict(extra="allow")

    symbol: str
    old_price: Optional[float | int | str] = None
    new_price: Optional[float | int | str] = None
    pct_change: Optional[float | int | str] = None
    report_date: Optional[str] = None


class QuotesPayload(BaseModel):
    """Schema for quotes payload."""
    quotes: list[QuoteItem]


class QuotesReceiveResponse(BaseModel):
    """Schema for quotes receive endpoint response."""
    status: str
    total_received: int


class QuotesStatusResponse(BaseModel):
    """Schema for quotes status endpoint response."""
    status: str
    total_quotes: int
    report_date: Optional[str] = None
    last_received_utc: Optional[str] = None
