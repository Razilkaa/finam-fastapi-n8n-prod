"""Text processing utility functions."""


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
