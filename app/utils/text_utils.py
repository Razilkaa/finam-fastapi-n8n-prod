"""Text processing utility functions."""
import re

from app.utils.constants import (
    MONTH_EN_TO_RU_GENITIVE,
    MONTH_EN_TO_RU_NOMINATIVE,
    QUARTER_EN_TO_RU,
)


def sanitize_text(text) -> str:
    """Санитизация текста для защиты от Excel formula injection."""
    if text is None or not isinstance(text, str):
        return ""
    text = text.strip()
    text = text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
    if text and text[0] in ('=', '+', '-', '@'):
        text = "'" + text
    return text


def has_cyrillic(text: str) -> bool:
    """Проверяет, содержит ли текст кириллицу."""
    if not text or not isinstance(text, str):
        return False
    return any('\u0400' <= char <= '\u04FF' for char in text)


def convert_month_suffix_to_ru(event_text: str) -> str:
    """
    Преобразует английские суффиксы месяцев/кварталов в конце события в русский формат.
    
    Примеры:
    - "Индекс цен на жилье S&P/CaseShiller (г/г) NOV" → "Индекс цен на жилье S&P/CaseShiller (г/г), ноябрь"
    - "Запасы сырой нефти от EIA JAN/23" → "Запасы сырой нефти от EIA, 23 января"
    """
    if not event_text or not isinstance(event_text, str):
        return event_text or ""
    
    text = event_text.strip()
    
    # Паттерн: MONTH/DAY в конце (например JAN/23, FEB/14)
    match_month_day = re.search(r'\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)/(\d{1,2})$', text, re.IGNORECASE)
    if match_month_day:
        month_en = match_month_day.group(1).upper()
        day = match_month_day.group(2)
        month_ru = MONTH_EN_TO_RU_GENITIVE.get(month_en, month_en)
        text = text[:match_month_day.start()].rstrip()
        return f"{text}, {day} {month_ru}"
    
    # Паттерн: только MONTH в конце (например NOV, DEC)
    match_month = re.search(r'\s+(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)$', text, re.IGNORECASE)
    if match_month:
        month_en = match_month.group(1).upper()
        month_ru = MONTH_EN_TO_RU_NOMINATIVE.get(month_en, month_en)
        text = text[:match_month.start()].rstrip()
        return f"{text}, {month_ru}"
    
    # Паттерн: квартал Q1-Q4 в конце
    match_quarter = re.search(r'\s+(Q[1-4])$', text, re.IGNORECASE)
    if match_quarter:
        quarter_en = match_quarter.group(1).upper()
        quarter_ru = QUARTER_EN_TO_RU.get(quarter_en, quarter_en)
        text = text[:match_quarter.start()].rstrip()
        return f"{text}, {quarter_ru}"
    
    return event_text
