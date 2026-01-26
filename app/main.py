"""Main application entry point."""
from fastapi import FastAPI

from app.core.config import APP_TITLE, APP_DESCRIPTION, APP_VERSION
from app.api.v1.endpoints import calendar

app = FastAPI(
    title=APP_TITLE,
    description=APP_DESCRIPTION,
    version=APP_VERSION
)

app.include_router(calendar.router)


@app.get("/")
async def root():
    """Root endpoint with API information."""
    return {
        "status": "ok",
        "message": "Calendar Generator API",
        "endpoints": {
            "POST /api/calendar/receive": "Приём данных от n8n (единый файл Events.json)",
            "GET /api/calendar/generate": "Генерация Excel файла",
            "GET /api/calendar/generate-word": "Генерация Word файла из шаблона",
            "GET /api/calendar/status": "Статус данных",
            "POST /api/calendar/clear": "Очистка данных",
        }
    }
