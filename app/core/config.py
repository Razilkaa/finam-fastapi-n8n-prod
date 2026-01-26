"""Application configuration."""
import os
from pathlib import Path

# Путь к шаблону Word
WORD_TEMPLATE_PATH = Path(os.getenv("WORD_TEMPLATE_PATH", "/app/Template.docx"))

# Настройки приложения
APP_TITLE = "Calendar Generator API"
APP_DESCRIPTION = "API для генерации экономического календаря"
APP_VERSION = "1.0.0"
