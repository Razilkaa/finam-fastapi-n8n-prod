# Calendar & Quotes API

FastAPI-приложение для:
- генерации экономического календаря в форматах Excel и Word;
- приёма котировок и формирования ежедневного Word-документа по шаблону.

Также поддерживается обновление Word-шаблонов через API без перезапуска сервиса.

## Структура проекта

```
main_prod/
├── app/
│   ├── __init__.py
│   ├── main.py                        # Точка входа приложения
│   ├── api/                           # API роутеры
│   │   ├── __init__.py
│   │   └── v1/
│   │       ├── __init__.py
│   │       └── endpoints/
│   │           ├── __init__.py
│   │           ├── calendar.py        # Эндпоинты календаря
│   │           ├── template.py        # Управление шаблоном календаря (Word)
│   │           └── quotes.py          # Эндпоинты котировок + шаблон котировок (Word)
│   ├── core/                          # Конфигурация
│   │   ├── __init__.py
│   │   └── config.py                  # Настройки приложения
│   ├── models/                        # Pydantic схемы
│   │   ├── __init__.py
│   │   └── schemas.py                 # Модели данных
│   ├── services/                      # Бизнес-логика
│   │   ├── __init__.py
│   │   ├── calendar_service.py        # Обработка данных календаря
│   │   ├── data_store.py              # Хранилище данных календаря
│   │   ├── excel_service.py           # Генерация Excel
│   │   ├── word_service.py            # Генерация Word календаря
│   │   ├── template_service.py        # Управление шаблоном календаря (runtime update)
│   │   ├── quotes_store.py            # Хранилище котировок
│   │   ├── quotes_doc_service.py      # Заполнение docx котировок по таблице
│   │   └── quotes_template_service.py # Управление шаблоном котировок (runtime update)
│   └── utils/                         # Утилиты
│       ├── __init__.py
│       ├── constants.py
│       ├── date_utils.py
│       └── text_utils.py
├── Template.docx                      # Шаблон Word (календарь)
├── Template_quotes.docx               # Шаблон Word (котировки)
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```

## Запуск

### Docker Compose

```bash
docker compose up --build
```

### Локально

#### С использованием uv (рекомендуется)

```bash
# Установка uv (если ещё не установлен)
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

- `GET /` — информация об API и список доступных ручек

### Календарь

- `POST /api/calendar/receive` — приём данных от n8n (единый массив `events`)
- `GET /api/calendar/status` — статус загруженных данных
- `GET /api/calendar/generate` — сгенерировать Excel
- `GET /api/calendar/generate-word` — сгенерировать Word по шаблону календаря
- `POST /api/calendar/clear` — очистить данные

### Шаблон календаря (Word)

- `GET /api/template` — информация о текущем шаблоне календаря
- `POST /api/template` — загрузить новый шаблон календаря (`.docx`)
- `GET /api/template/download` — скачать текущий шаблон календаря (`.docx`)

### Котировки

- `POST /api/quotes/receive` — приём котировок (поддерживает `{ "quotes": [...] }` или `[...]`)
- `GET /api/quotes/status` — статус котировок
- `GET /api/quotes/daily/word` — сформировать Word-документ котировок по шаблону

### Шаблон котировок (Word)

- `GET /api/quotes/template` — информация о текущем шаблоне котировок
- `POST /api/quotes/template` — загрузить новый шаблон котировок (`.docx`)
- `GET /api/quotes/template/download` — скачать текущий шаблон котировок (`.docx`)

## Переменные окружения

- `WORD_TEMPLATE_PATH` — путь к шаблону календаря (по умолчанию: `/app/Template.docx`)
- `QUOTES_TEMPLATE_PATH` — путь к шаблону котировок (по умолчанию: `/app/Template_quotes.docx`)
