# file: remove_duplicates_wb.py

import os
import xlwings as xw

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')
SHEET_FACTS = 'ФинотчетыWB'

def get_workbook():
    try:
        wb = xw.Book.caller()
        app = None
    except:
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_PATH)
    return wb, app

def main():
    wb, app = get_workbook()
    ws = wb.sheets[SHEET_FACTS]
    data = ws.range('A1').expand().value
    if not data or len(data) < 2:
        print("Нет данных для обработки")
        return

    header = data[0]
    rows = data[1:]
    idx = {k: i for i, k in enumerate(header)}

    seen_keys = set()
    unique_rows = []

    for row in rows:
        if not row or len(row) < len(header):
            continue
        key = (
            str(row[idx['Организация']]).strip(),
            str(row[idx['Номер_отчёта']]).strip(),
            str(row[idx['Артикул_WB']]).strip().split('.')[0],
            str(row[idx['Артикул_продавца']]).strip()
        )
        if key in seen_keys:
            continue
        seen_keys.add(key)
        unique_rows.append(row)

    # Очистка и перезапись
    ws.clear_contents()
    ws.range('A1').value = [header] + unique_rows
    print(f"Оставлено уникальных строк: {len(unique_rows)}")

    if app:
        wb.save()
        app.quit()

if __name__ == '__main__':
    main()
