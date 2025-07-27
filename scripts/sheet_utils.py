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
    "КомиссияWB": {"color": "#92D050", "position": 30},
    "План_Продаж": {"color": "#FFC000", "position": 12},
    # остальные листы…
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
