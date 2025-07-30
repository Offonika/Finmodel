"""Utilities for working with Excel sheets."""

import xlwings as xw


def hex_to_excel_tab_color(hex_str: str) -> int:
    """Convert ``#RRGGBB`` HEX string to Excel's BGR tab color integer."""
    hex_clean = hex_str.strip().lstrip("#")
    if len(hex_clean) != 6:
        raise ValueError("HEX color must be in format #RRGGBB")
    r = int(hex_clean[0:2], 16)
    g = int(hex_clean[2:4], 16)
    b = int(hex_clean[4:6], 16)
    return b + g * 256 + r * 256 * 256


SHEET_SETTINGS: dict[str, dict[str, int | str]] = {
    # 📋 Меню и системные
    "Меню": {"color": "#D9D9D9", "position": 0},

    # ⚙ Настройки
    "Настройки": {"color": "#BDD7EE", "position": 1},
    "НастройкиОрганизаций": {"color": "#BDD7EE", "position": 2},

    # 📊 Дашборд
    "Дашборд": {"color": "#F4B084", "position": 5},

    # 🧾 Плановые расчёты
    "РасчетПлановыхПоказателей": {"color": "#D9D9D9", "position": 8},
    "РасчетЭкономикиWB": {"color": "#92D050", "position": 14},
    "РасчетЭкономикиОзон": {"color": "#92D050", "position": 15},
    "РасчетСебестоимости": {"color": "#92D050", "position": 18},
    "РасчетЗарплаты": {"color": "#92D050", "position": 25},
    # 💵 Планирование
    "ПланПродажWB": {"color": "#FFD966", "position": 30},
    "ПланПродажОзон": {"color": "#FFD966", "position": 35},
    "ПланВыручкиWB": {"color": "#FFD966", "position": 40},
    "ПланВыручкиОзон": {"color": "#FFD966", "position": 45},

    # 📈 Финотчеты
    "ФинотчетыWB": {"color": "#C6E0B4", "position": 50},
    "ФинотчетыОзон": {"color": "#C6E0B4", "position": 55},

    # 💼 Зарплата и расходы
    "Зарплата": {"color": "#B4C6E7", "position": 60},
    "ПрочиеРасходы": {"color": "#B4C6E7", "position": 63},

    # 🗓 Сезонность / справочники
    "Сезонность": {"color": "#DEEAF6", "position": 67},
    "ЗакупочныеЦены": {"color": "#DEEAF6", "position": 70},
    "ТаможенныеПошлины": {"color": "#DEEAF6", "position": 75},
    "Номенклатура_WB": {"color": "#DEEAF6", "position": 80},

    # 📦 Цены
    "ЦеныWB": {"color": "#FBE4D5", "position": 83},
    "ЦеныОзон": {"color": "#FBE4D5", "position": 85},

    # 📋 Комиссия и начисления
    "КомиссияWB": {"color": "#C6E0B4", "position": 88},
    "НачисленияУслугОзон": {"color": "#C6E0B4", "position": 90},

    # 📊 Показатели
    "Показатели": {"color": "#F4B084", "position": 95}
}



def apply_sheet_settings(wb: xw.Book, sheet_name: str) -> None:
    """Apply color and position for ``sheet_name`` according to ``SHEET_SETTINGS``."""
    if sheet_name not in [s.name for s in wb.sheets]:
        return

    sheet = wb.sheets[sheet_name]
    cfg = SHEET_SETTINGS.get(sheet_name)
    if not cfg:
        return

    color = cfg.get("color")
    if color:
        try:
            sheet.api.Tab.Color = hex_to_excel_tab_color(str(color))
        except Exception:
            pass

    pos = cfg.get("position")
    if isinstance(pos, int) and pos > 0:
        try:
            target = min(pos, len(wb.sheets))
            if sheet.index != target:
                if target == len(wb.sheets):
                    sheet.api.Move(After=wb.sheets[target - 1].api)
                else:
                    sheet.api.Move(Before=wb.sheets[target - 1].api)
        except Exception:
            pass
