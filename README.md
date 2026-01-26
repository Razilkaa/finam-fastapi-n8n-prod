# Calendar Generator API

FastAPI приложение для генерации экономического календаря в форматах Excel и Word.

## Структура проекта

```
main_prod/
├── app/
│   ├── __init__.py
│   ├── main.py                 # Точка входа приложения
│   ├── api/                    # API роутеры
│   │   ├── __init__.py
│   │   └── v1/
│   │       ├── __init__.py
│   │       └── endpoints/
│   │           ├── __init__.py
│   │           └── calendar.py  # Эндпоинты календаря
│   ├── core/                   # Конфигурация
│   │   ├── __init__.py
│   │   └── config.py           # Настройки приложения
│   ├── models/                 # Pydantic схемы
│   │   ├── __init__.py
│   │   └── schemas.py          # Модели данных
│   ├── services/               # Бизнес-логика
│   │   ├── __init__.py
│   │   ├── calendar_service.py # Обработка данных календаря
│   │   ├── data_store.py       # Хранилище данных
│   │   ├── excel_service.py    # Генерация Excel
│   │   └── word_service.py     # Генерация Word
│   └── utils/                  # Утилиты
│       ├── __init__.py
│       ├── constants.py        # Константы
│       ├── date_utils.py       # Работа с датами
│       └── text_utils.py       # Обработка текста
├── Template.docx               # Шаблон Word документа
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```

## Запуск

### Docker Compose

```bash
docker-compose up --build
```

### Локально

#### С использованием uv (рекомендуется)

```bash
# Установка uv (если еще не установлен)
# Windows: powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
# Linux/Mac: curl -LsSf https://astral.sh/uv/install.sh | sh

# Установка зависимостей
uv pip install -r requirements.txt

# Запуск приложения
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

#### С использованием pip (альтернатива)

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

## API Endpoints

- `GET /` - Информация об API
- `POST /api/calendar/receive` - Приём данных от n8n
- `GET /api/calendar/status` - Статус данных
- `GET /api/calendar/generate` - Генерация Excel файла
- `GET /api/calendar/generate-word` - Генерация Word файла
- `POST /api/calendar/clear` - Очистка данных

## Переменные окружения

- `WORD_TEMPLATE_PATH` - Путь к шаблону Word (по умолчанию: `/app/Template.docx`)
