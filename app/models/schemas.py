"""Pydantic schemas for API requests and responses."""
from typing import Optional
from pydantic import BaseModel


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
