"""Date and time utility functions."""
from datetime import date, datetime, timedelta
from typing import Optional


def parse_date(date_str) -> Optional[date]:
    """Парсинг даты. Защита от None, неверных типов и форматов."""
    if date_str is None or not isinstance(date_str, str):
        return None
    date_str = date_str.strip()
    if not date_str:
        return None
    try:
        if "-" in date_str:
            return datetime.strptime(date_str[:10], "%Y-%m-%d").date()
        if "." in date_str:
            return datetime.strptime(date_str[:10], "%d.%m.%Y").date()
    except ValueError:
        pass
    return None


def parse_time_for_sort(time_str) -> tuple[int, int]:
    """Парсинг времени для сортировки. Защита от None и неверных типов."""
    if time_str is None or not isinstance(time_str, str):
        return (99, 99)
    time_str = time_str.strip()
    if not time_str:
        return (99, 99)
    time_str = time_str.upper()
    is_pm = "PM" in time_str
    is_am = "AM" in time_str
    time_clean = time_str.replace("AM", "").replace("PM", "").strip()
    try:
        if ":" in time_clean:
            parts = time_clean.split(":")
            hours = int(parts[0])
            minutes = int(parts[1]) if len(parts) > 1 else 0
            if is_pm and hours != 12:
                hours += 12
            elif is_am and hours == 12:
                hours = 0
            return (hours, minutes)
    except (ValueError, IndexError):
        pass
    return (99, 99)


def format_time_display(time_str) -> str:
    """Возвращает время как есть из исходного JSON (с AM/PM). Защита от None."""
    if time_str is None or not isinstance(time_str, str):
        return ""
    return time_str.strip()


def get_monday_of_week(d: date) -> date:
    """Получить понедельник недели."""
    return d - timedelta(days=d.weekday())


def get_week_dates(monday: date) -> list[date]:
    """Получить рабочие дни недели (пн-пт)."""
    return [monday + timedelta(days=i) for i in range(5)]


def format_date_ru(d: date) -> str:
    """Форматирование даты для русского языка."""
    return f"{d.day:02d}.{d.month:02d}.{d.year}"


def format_date_en(d: date) -> str:
    """Форматирование даты для английского языка."""
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    months = ["January", "February", "March", "April", "May", "June",
              "July", "August", "September", "October", "November", "December"]
    return f"{days[d.weekday()]} {months[d.month - 1]} {d.day} {d.year}"


def format_sheet_name_ru(d: date) -> str:
    """Форматирование имени листа для русского языка."""
    return f"Календарь {d.day:02d}.{d.month:02d}.{d.year}"


def format_sheet_name_en(d: date) -> str:
    """Форматирование имени листа для английского языка."""
    return f"Economic calendar_{d.day:02d}.{d.month:02d}.{str(d.year)[2:]}"


def group_items_by_date(items: list[dict]) -> dict[date, list[dict]]:
    """Группировка элементов по дате."""
    grouped: dict[date, list[dict]] = {}
    for item in items:
        if not isinstance(item, dict):
            continue
        d = parse_date(item.get("date", ""))
        if d:
            grouped.setdefault(d, []).append(item)
    return grouped


def choose_reference_monday(
    events_by_date: dict[date, list[dict]],
    holidays_by_date: dict[date, list[dict]],
) -> date:
    """Choose the most likely week to render, robust against outlier dates."""
    week_scores: dict[date, int] = {}

    for d, items in events_by_date.items():
        if d.weekday() >= 5:
            continue
        monday = get_monday_of_week(d)
        week_scores[monday] = week_scores.get(monday, 0) + len(items)

    for d, items in holidays_by_date.items():
        if d.weekday() >= 5:
            continue
        monday = get_monday_of_week(d)
        week_scores[monday] = week_scores.get(monday, 0) + len(items)

    if not week_scores:
        return get_monday_of_week(date.today())

    today_monday = get_monday_of_week(date.today())
    best_monday, _score = max(
        week_scores.items(),
        key=lambda kv: (kv[1], -abs((kv[0] - today_monday).days)),
    )
    return best_monday
