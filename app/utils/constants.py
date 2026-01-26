"""Application constants."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Цвета
COLOR_RED = "FF0000"
COLOR_BLACK = "333333"
COLOR_DATE_TEXT = "212529"

# Заливки
GRAY_FILL = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
NO_FILL = PatternFill(fill_type=None)

# Выравнивание
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")

# Границы
BORDER_THIN = Side(style="thin", color="000000")
BORDER_MEDIUM = Side(style="medium", color="000000")

# Шрифты
FONT_HEADER = Font(name="Arial", size=11, bold=True)
FONT_DATE = Font(name="Arial", size=11, bold=False, color=COLOR_DATE_TEXT)
FONT_TIME = Font(name="Arial", size=11, bold=True, color=COLOR_BLACK)
FONT_TIME_RED = Font(name="Arial", size=11, bold=True, color=COLOR_RED)
FONT_EVENT = Font(name="Arial", size=11, bold=False, color=COLOR_BLACK)
FONT_EVENT_RED = Font(name="Arial", size=11, bold=False, color=COLOR_RED)
FONT_HOLIDAY = Font(name="Arial", size=11, bold=False, color=COLOR_RED)

# Названия стран
COUNTRY_NAMES_EN = {
    "US": "US", "GB": "UK", "EU": "EU", "EA": "EU",
    "DE": "Germany", "JP": "Japan", "CN": "China", "CH": "Switzerland",
}
COUNTRY_NAMES_RU = {
    "US": "США", "GB": "Великобритания", "EU": "ЕС", "EA": "ЕС",
    "DE": "Германия", "JP": "Япония", "CN": "Китай", "CH": "Швейцария",
}
