# calculate_avg_logistics.py

from pathlib import Path
import xlwings as xw
import pandas as pd

EXCEL_PATH = Path(__file__).resolve().parents[1] / 'Finmodel.xlsm'

SHEET_SOURCE = 'НачисленияУслугОзон'
SHEET_OUT    = 'Показатели'
TABLE_NAME   = 'AvgLogisticsTable'
TABLE_STYLE  = 'TableStyleMedium7'

def get_workbook():
    try:
        wb = xw.Book.caller()
        app = None
        print('→ Запуск из Excel-макроса')
    except Exception:
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_PATH)
        print(f'→ Запуск из терминала, открыт файл: {EXCEL_PATH}')
    return wb, app

def safe_float(val):
    try:
        if pd.isna(val):
            return 0.0
        return float(str(val).replace(',', '.').replace('\xa0', '').replace(' ', ''))
    except Exception:
        return 0.0

def main():
    print('=== Старт calculate_avg_logistics ===')
    wb, app = get_workbook()

    try:
        ws = wb.sheets[SHEET_SOURCE]
        df = ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
        print(f'→ Прочитано строк из {SHEET_SOURCE}: {len(df)}')
    except Exception as e:
        print(f'❌ Нет листа {SHEET_SOURCE}: {e}')
        if app:
            app.quit()
        return

    # Проверяем нужные колонки
    required = [
        "ПроданоШт","Логистика","Сборка заказа","Обработка отправления","Магистраль",
        "Последняя миля","Обратная магистраль","Обработка возврата","Обратная логистика",
        "Оплата эквайринга","ВыручкаБезСкидок"
    ]
    for c in required:
        if c not in df.columns:
            print(f'❌ Нет колонки {c} в {SHEET_SOURCE}')
            if app:
                app.quit()
            return

    log_fields = [
        "Сборка заказа",
        "Обработка отправления",
        "Магистраль",
        "Последняя миля",
        "Обратная магистраль",
        "Обработка возврата",
        "Обратная логистика"
    ]

    total_qty = 0
    total_log = 0
    total = {fld: 0.0 for fld in log_fields}
    ekv_sum = 0.0
    rev_sum = 0.0

    for _, row in df.iterrows():
        qty = safe_float(row["ПроданоШт"])
        full_log = safe_float(row["Логистика"])
        if qty == 0:
            continue
        total_qty += qty
        total_log += full_log

        for f in log_fields:
            total[f] += safe_float(row[f])

        # Эквайринг и выручка
        ekv = safe_float(row["Оплата эквайринга"])
        rev = safe_float(row["ВыручкаБезСкидок"])
        if rev > 0 and ekv >= 0:
            ekv_sum += ekv
            rev_sum += rev

    avg_ekv_percent = round((ekv_sum / rev_sum) * 100, 2) if rev_sum > 0 else 0.0

    # Заголовки для Excel
    header = [
        "Сборка заказа, ₽",
        "Обработка отправления, ₽",
        "Магистраль, ₽",
        "Последняя миля, ₽",
        "Обратная магистраль, ₽",
        "Обработка возврата, ₽",
        "Обратная логистика, ₽",
        "Логистика, ₽",
        "Эквайринг, %"
    ]
    row_out = []
    for f in log_fields:
        avg = round(abs(total[f]) / total_qty, 2) if total_qty else 0.0
        row_out.append(avg)
    avg_full_log = round(abs(total_log) / total_qty, 2) if total_qty else 0.0
    row_out.append(avg_full_log)
    row_out.append(avg_ekv_percent)

    # Запись в Excel
    try:
        out_ws = wb.sheets[SHEET_OUT]
        out_ws.clear()
        print(f'→ Лист {SHEET_OUT} очищен')
    except Exception:
        out_ws = wb.sheets.add(SHEET_OUT)
        print(f'→ Лист {SHEET_OUT} создан')

    out_ws.range(1,1).value = header
    out_ws.range(2,1).value = row_out

    # Форматировать как умную таблицу (TableStyleMedium7)
    for tbl in out_ws.tables:
        if tbl.name == TABLE_NAME:
            tbl.delete()
    rng = out_ws.range((1,1), (2, len(header)))
    out_ws.tables.add(rng, name=TABLE_NAME, table_style_name=TABLE_STYLE, has_headers=True)
    out_ws.range('A1').expand().columns.autofit()
    out_ws.api.Rows(1).Font.Bold = True

    print('=== Скрипт успешно завершён ===')
    if app:
        wb.save()
        app.quit()

if __name__ == '__main__':
    main()
