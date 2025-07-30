import xlwings as xw

TABLE_STYLE = 'TableStyleMedium7'


def format_table(ws: xw.main.Sheet, table_range: xw.main.Range, table_name: str,
                 table_style: str = TABLE_STYLE) -> None:
    """Create or replace a table and apply common formatting."""
    try:
        for tbl in ws.tables:
            if tbl.name == table_name:
                tbl.delete()
    except Exception:
        pass
    ws.tables.add(table_range, name=table_name, table_style_name=table_style,
                  has_headers=True)
    try:
        ws.range('A1').expand().columns.autofit()
    except Exception:
        try:
            ws.autofit()
        except Exception:
            pass
    try:
        ws.api.Rows(1).Font.Bold = True
    except Exception:
        pass


def autofit_safe(ws: xw.main.Sheet) -> None:
    """Safely autofit all columns of a sheet."""
    try:
        ws.autofit()
    except Exception:
        try:
            ws.range('A1').expand().columns.autofit()
        except Exception:
            pass
