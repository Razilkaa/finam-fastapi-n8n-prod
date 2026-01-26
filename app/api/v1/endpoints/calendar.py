"""Calendar API endpoints."""
from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from app.models.schemas import EventsPayload, StatusResponse, ReceiveResponse
from app.services.data_store import data_store
from app.services.calendar_service import split_events_data
from app.services.excel_service import generate_excel
from app.services.word_service import generate_word, get_output_filename
from app.core.config import WORD_TEMPLATE_PATH
from app.utils.date_utils import group_items_by_date, choose_reference_monday

router = APIRouter(prefix="/api/calendar", tags=["calendar"])


@router.post("/receive", response_model=ReceiveResponse)
async def receive_data(payload: EventsPayload):
    """Приём единого массива событий. Автоматически разделяет по языку и типу."""
    all_events = [ev.model_dump() for ev in payload.events]
    
    work_en, work_ru, holidays_en, holidays_ru = split_events_data(all_events)
    
    data_store["work_en"] = work_en
    data_store["work_ru"] = work_ru
    data_store["holidays_en"] = holidays_en
    data_store["holidays_ru"] = holidays_ru
    
    return ReceiveResponse(
        status="ok",
        total_received=len(all_events),
        split={
            "work_en": len(work_en),
            "work_ru": len(work_ru),
            "holidays_en": len(holidays_en),
            "holidays_ru": len(holidays_ru),
        }
    )


@router.get("/status", response_model=StatusResponse)
async def get_status():
    """Получить статус данных."""
    return StatusResponse(
        status="ok",
        data={key: len(value) for key, value in data_store.items()}
    )


@router.get("/generate")
async def generate_calendar():
    """Генерация Excel файла."""
    try:
        buffer = generate_excel(
            work_en=data_store["work_en"],
            work_ru=data_store["work_ru"],
            holidays_en=data_store["holidays_en"],
            holidays_ru=data_store["holidays_ru"],
        )
        
        combined_events_by_date = group_items_by_date(data_store["work_en"] + data_store["work_ru"])
        combined_holidays_by_date = group_items_by_date(data_store["holidays_en"] + data_store["holidays_ru"])
        
        monday = choose_reference_monday(combined_events_by_date, combined_holidays_by_date)
        filename = f"Calendar_{monday.day:02d}.{monday.month:02d}.{monday.year}.xlsx"
        
        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating Excel: {str(e)}")


@router.get("/generate-word")
async def generate_word_calendar():
    """Генерация Word документа из шаблона."""
    try:
        if not WORD_TEMPLATE_PATH.exists():
            raise HTTPException(
                status_code=404,
                detail=f"Word template not found: {WORD_TEMPLATE_PATH}. "
                       f"Mount Template.docx to /app/Template.docx or set WORD_TEMPLATE_PATH env var."
            )
        
        buffer = generate_word(
            work_en=data_store["work_en"],
            work_ru=data_store["work_ru"],
            holidays_en=data_store["holidays_en"],
            holidays_ru=data_store["holidays_ru"],
            template_path=WORD_TEMPLATE_PATH,
        )
        
        filename = get_output_filename(
            events=(data_store["work_ru"] + data_store["work_en"]),
            holidays=(data_store["holidays_ru"] + data_store["holidays_en"]),
        )
        
        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating Word: {str(e)}")


@router.post("/clear")
async def clear_data():
    """Очистка данных."""
    for key in data_store:
        data_store[key] = []
    return {"status": "ok", "message": "Data cleared"}
