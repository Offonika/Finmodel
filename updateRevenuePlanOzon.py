# updaterevenueplan_ozon.py

import os
import xlwings as xw
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')

SHEET_SALES = 'ПланПродажОзон'
SHEET_OUT   = 'ПланВыручкиОзон'
TABLE_NAME  = 'PlanOzonRevenueTable'
TABLE_STYLE = 'TableStyleMedium7'

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

def main():
    print("=== Старт update_revenue_plan_ozon ===")
    wb, app = get_workbook()

    try:
        sales_ws = wb.sheets[SHEET_SALES]
        df = sales_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
        print(f"→ Прочитано строк из {SHEET_SALES}: {len(df)}")
    except Exception as e:
        print(f"❌ Нет листа {SHEET_SALES}: {e}")
        if app: app.quit()
        return

    df = df[df['Организация'].str.lower() != 'итого']  # убираем строки "Итого", если есть

    # Найти все месячные колонки
    month_cols = [col for col in df.columns if col.startswith('Мес.')]
    print(f"→ Месячные столбцы: {month_cols}")
    if not month_cols:
        print("❌ Нет месячных колонок!")
        if app: app.quit()
        return

    # Проверка нужных колонок
    for c in ['Организация', 'Артикул_поставщика', 'SKU', 'Плановая цена']:
        if c not in df.columns:
            print(f"❌ Нет колонки {c} в {SHEET_SALES}")
            if app: app.quit()
            return

    # Вычислить выручку по месяцам и итог
    res_rows = []
    for i, row in df.iterrows():
        org = row['Организация']
        art = row['Артикул_поставщика']
        sku = row['SKU']
        price = float(str(row['Плановая цена']).replace(',', '.').replace('₽', '').replace(' ', '').replace(' ','') or 0)
        revs = [float(row[col] or 0) * price for col in month_cols]
        total = sum(revs)
        res_rows.append([org, art, sku] + revs + [total])
    print(f"→ Итоговых строк: {len(res_rows)}")
    header = ['Организация', 'Артикул_поставщика', 'SKU'] + month_cols + ['Всего']

    # --- Вывод ---
    # --- Вывод ---
    try:
        out_ws = wb.sheets[SHEET_OUT]
        out_ws.clear()
        print(f'→ Лист {SHEET_OUT} очищен')
    except:
        out_ws = wb.sheets.add(SHEET_OUT)
        print(f'→ Лист {SHEET_OUT} создан')

    # Установка цвета ярлыка #FFC000 (золотой)
    out_ws.range(1,1).value = header
    if res_rows:
        out_ws.range(2,1).value = res_rows

    # Цвет ярлыка #FFC000 (золотой)
        # Цвет ярлыка #FFC000 (золотой)
    try:
        out_ws.api.Tab.Color = 0x00C0FF  # правильный золотой #FFC000 (BGR)
        print("→ Цвет ярлыка #FFC000 установлен")
    except Exception as e:
        print(f"⚠️ Не удалось установить цвет ярлыка: {e}")

    # Переместить лист на позицию 9 (если надо)
    try:
        if out_ws.index != 10  and  len(wb.sheets) >= 9:
            before_sheet = wb.sheets[8]
            out_ws.api.Move(Before=before_sheet.api)
            print("→ Лист перемещён на позицию 9")
    except Exception as e:
        print(f"⚠️ Не удалось переместить лист: {e}")



    # Итоговая строка (сумма по каждому месяцу и по "Всего")
    last_row = len(res_rows) + 2
    total_row = ['Итого', '', '']
    for j in range(3, len(header)):
        col_letter = xw.utils.col_name(j+1)
        total_row.append(f'=SUM({col_letter}2:{col_letter}{last_row-1})')
    out_ws.range(last_row, 1).value = total_row

    # Форматировать как умную таблицу TableStyleMedium7
    for tbl in out_ws.tables:
        if tbl.name == TABLE_NAME:
            tbl.delete()
    tbl_range = out_ws.range((1,1), (last_row, len(header)))
    out_ws.tables.add(tbl_range, name=TABLE_NAME, table_style_name=TABLE_STYLE, has_headers=True)

    out_ws.range('A1').expand().columns.autofit()
    out_ws.api.Rows(1).Font.Bold = True
    out_ws.api.Application.ActiveWindow.SplitRow = 1
    out_ws.api.Application.ActiveWindow.FreezePanes = True

    print("=== Скрипт успешно завершён ===")
    if app: wb.save(); app.quit()

if __name__ == '__main__':
    main()
