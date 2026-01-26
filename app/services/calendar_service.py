"""Calendar data processing service."""
from app.utils.text_utils import has_cyrillic


def split_events_data(all_data: list) -> tuple[list, list, list, list]:
    """Разделяет данные из единого файла на 4 списка: work_en, work_ru, holidays_en, holidays_ru."""
    work_en = []
    work_ru = []
    holidays_en = []
    holidays_ru = []
    
    for item in all_data:
        if not isinstance(item, dict):
            continue
        
        is_holiday = "holiday" in item and item.get("holiday")
        text_to_check = item.get("holiday", "") if is_holiday else item.get("event", "")
        is_russian = has_cyrillic(text_to_check)
        
        if is_holiday:
            if is_russian:
                holidays_ru.append(item)
            else:
                holidays_en.append(item)
        else:
            if is_russian:
                work_ru.append(item)
            else:
                work_en.append(item)
    
    return work_en, work_ru, holidays_en, holidays_ru
