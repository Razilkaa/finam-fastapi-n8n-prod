"""Quotes API endpoints."""

from __future__ import annotations

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse

from app.models.schemas import (
    QuoteItem,
    QuotesPayload,
    QuotesReceiveResponse,
    QuotesStatusResponse,
)
from app.services.quotes_doc_service import fill_template, get_quotes_filename, parse_quotes
from app.services.quotes_store import quotes_store, set_quotes
from app.services.quotes_template_service import (
    DOCX_MIME,
    get_template_info,
    get_template_path,
    update_template_bytes,
)

router = APIRouter(prefix="/api/quotes", tags=["quotes"])


@router.post("/receive", response_model=QuotesReceiveResponse)
async def receive_quotes(payload: QuotesPayload | list[QuoteItem]):
    """Receive quotes JSON (either {quotes:[...]} or a raw list)."""
    items = payload.quotes if isinstance(payload, QuotesPayload) else payload
    raw_quotes = [q.model_dump() for q in items]

    _quotes, report_dt = parse_quotes(raw_quotes)
    report_date_str = report_dt.isoformat() if report_dt is not None else None

    set_quotes(quotes=raw_quotes, report_date=report_date_str)

    return QuotesReceiveResponse(status="ok", total_received=len(items))


@router.get("/status", response_model=QuotesStatusResponse)
async def quotes_status():
    """Get current quotes status."""
    return QuotesStatusResponse(
        status="ok",
        total_quotes=len(quotes_store["quotes"]),
        report_date=quotes_store["report_date"],
        last_received_utc=quotes_store["last_received_utc"],
    )


@router.get("/daily/word")
async def daily_quotes_word():
    """Generate a Word document using the current quotes and the quotes template."""
    if not quotes_store["quotes"]:
        raise HTTPException(status_code=400, detail="No quotes received yet.")

    try:
        template_path = get_template_path()
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))

    try:
        quotes, report_dt = parse_quotes(quotes_store["quotes"])
        buffer, updated_rows = fill_template(template_path=template_path, quotes=quotes)
        filename = get_quotes_filename(report_dt)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating Word: {str(e)}")

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f"attachment; filename={filename}",
            "X-Updated-Rows": str(updated_rows),
        },
    )


@router.get("/template")
async def quotes_template_info():
    """Get current quotes template metadata."""
    try:
        return {"status": "ok", "template": get_template_info()}
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))


@router.get("/template/download")
async def download_quotes_template():
    """Download current quotes template file."""
    try:
        path = get_template_path()
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))

    return FileResponse(
        path=path,
        media_type=DOCX_MIME,
        filename=path.name,
    )


@router.post("/template")
async def upload_quotes_template(file: UploadFile = File(...)):
    """Upload and activate a new quotes Word template (.docx) without restarting the service."""
    data = await file.read()
    try:
        info = update_template_bytes(
            data,
            filename=file.filename,
            content_type=file.content_type,
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except OSError as e:
        raise HTTPException(status_code=500, detail=f"Failed to save template: {e}")

    return {"status": "ok", "template": info}
