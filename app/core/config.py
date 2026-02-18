"""Application configuration."""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]

# Путь к шаблону Word
WORD_TEMPLATE_FALLBACK_PATH = BASE_DIR / "Template.docx"
_default_template_path = (
    str(WORD_TEMPLATE_FALLBACK_PATH) if os.name == "nt" else "/app/Template.docx"
)
WORD_TEMPLATE_PATH = Path(os.getenv("WORD_TEMPLATE_PATH", _default_template_path))

# Quotes Word template
QUOTES_TEMPLATE_FALLBACK_PATH = BASE_DIR / "Template_quotes.docx"
_default_quotes_template_path = (
    str(QUOTES_TEMPLATE_FALLBACK_PATH) if os.name == "nt" else "/app/Template_quotes.docx"
)
QUOTES_TEMPLATE_PATH = Path(os.getenv("QUOTES_TEMPLATE_PATH", _default_quotes_template_path))

# Настройки приложения
APP_TITLE = "Calendar Generator API"
APP_DESCRIPTION = "API для генерации экономического календаря"
APP_VERSION = "1.0.0"
