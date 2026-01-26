"""Data storage service."""
from typing import TypedDict


class DataStore(TypedDict):
    """Type definition for data store."""
    work_en: list[dict]
    work_ru: list[dict]
    holidays_en: list[dict]
    holidays_ru: list[dict]


# Глобальное хранилище данных
data_store: DataStore = {
    "work_en": [],
    "work_ru": [],
    "holidays_en": [],
    "holidays_ru": [],
}
