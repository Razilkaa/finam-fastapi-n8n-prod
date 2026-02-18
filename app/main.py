"""Main application entry point."""
from fastapi import FastAPI

from app.core.config import APP_TITLE, APP_DESCRIPTION, APP_VERSION
from app.api.v1.endpoints import calendar
from app.api.v1.endpoints import template
from app.api.v1.endpoints import quotes

app = FastAPI(
    title=APP_TITLE,
    description=APP_DESCRIPTION,
    version=APP_VERSION
)

app.include_router(calendar.router)
app.include_router(template.router)
app.include_router(quotes.router)


@app.get("/")
async def root():
    """Root endpoint with API information."""
    return {
        "status": "ok",
        "message": "Excel/Word Generator API",
        "endpoints": {
            "POST /api/calendar/receive": "Приём данных от n8n (единый файл Events.json)",
            "GET /api/calendar/generate": "Генерация Excel файла",
            "GET /api/calendar/generate-word": "Генерация Word файла из шаблона",
            "GET /api/calendar/status": "Статус данных",
            "POST /api/calendar/clear": "Очистка данных",
            "GET /api/template": "Информация о шаблоне календаря",
            "POST /api/template": "Загрузить новый шаблон календаря (.docx)",
            "GET /api/template/download": "Скачать текущий шаблон календаря (.docx)",
            "GET /api/quotes/status": "Статус котировок",
            "POST /api/quotes/receive": "Приём котировок (JSON)",
            "GET /api/quotes/daily/word": "Сформировать Word-документ с котировками",
            "GET /api/quotes/template": "Информация о шаблоне котировок",
            "POST /api/quotes/template": "Загрузить новый шаблон котировок (.docx)",
            "GET /api/quotes/template/download": "Скачать текущий шаблон котировок (.docx)",
        }
    }
