# update_revenue_plan.py

import os
import xlwings as xw
import pandas as pd
import re

from scripts.style_utils import format_table

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'excel', 'Finmodel.xlsm') 


SHEET_SALES   = 'План_ПродажWB'
SHEET_REVENUE = 'План_ВыручкиWB'
TABLE_NAME    = 'RevenuePlanTable'
TABLE_STYLE   = 'TableStyleMedium7'

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
    if pd.isna(val):
        return 0.0
    try:
        return float(str(val).replace(',', '.').replace(' ','').replace(' ','')) # Убирает пробелы и неразрывные
    except Exception:
        return 0.0

def main():
    print("=== Старт update_revenue_plan ===")
    wb, app = get_workbook()
    print('→ Открыт файл:', wb.fullname)
    print('→ Листы книги:', [s.name for s in wb.sheets])

    # 1. Чтение исходного плана продаж
    try:
        sales_ws = wb.sheets[SHEET_SALES]
        df = sales_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
        print(f'→ Лист {SHEET_SALES} считан, строк: {len(df)}')
    except Exception as e:
        print(f'❌ Ошибка: не найден лист {SHEET_SALES} или не удалось считать данные\n{e}')
        if app: app.quit()
        return

    # 2. Убираем строки "Итого"
    df = df[df['Организация'].astype(str).str.lower() != 'итого']
    print('→ После фильтра "Итого" строк:', len(df))

    # 3. Месячные столбцы
    hdr = list(df.columns)
    price_col = 'Плановая цена, ₽'
    month_cols = [h for h in hdr if re.match(r'^Мес\.\d{2}$', str(h).strip())]
    print('→ Месячные столбцы:', month_cols)
    print('→ Заголовки:', hdr)
    if len(df) > 0:
        print('→ Первая строка:', df.iloc[0].to_dict())

    # 4. Итоговая таблица: выручка по месяцам = продажи × цена
    rows = []
    log_count = 0
    for idx, r in df.iterrows():
        org    = r['Организация']
        vendor = r['Артикул_поставщика']
        subj   = r['Предмет']
        price  = safe_float(r[price_col])
        sales_by_month = [safe_float(r[m]) for m in month_cols]
        revs   = [s * price for s in sales_by_month]
        total  = sum(revs)
        if log_count < 5:
            print(f'[{idx}] {org}, {vendor}: цена={price}, продажи={sales_by_month}, выручка={revs}, всего={total}')
        log_count += 1
        if any(revs):
            rows.append([org, vendor, subj] + revs + [total])
    print(f'→ Итог: записано строк в план выручки: {len(rows)}')

    rev_hdr = ['Организация', 'Артикул_поставщика', 'Предмет'] + month_cols + ['Всего']

    # 5. Запись на лист "План_ВыручкиWB"
    try:
        rev_ws = wb.sheets[SHEET_REVENUE]
        rev_ws.clear()
        print(f'→ Лист {SHEET_REVENUE} очищен')
    except:
        rev_ws = wb.sheets.add(SHEET_REVENUE)
        print(f'→ Лист {SHEET_REVENUE} создан')
    # --- Цвет ярлыка и позиция листа ---
    try:
        rev_ws.api.Tab.Color = 0x00C0FF  # золотой #FFC000 (BGR)
        print("→ Цвет ярлыка #FFC000 установлен")
    except Exception as e:
        print(f"⚠️ Не удалось установить цвет ярлыка: {e}")

    try:
        sheet_count = len(wb.sheets)
        pos = 14 if sheet_count >= 15 else sheet_count
        if rev_ws.index != pos:
            rev_ws.api.Move(Before=wb.sheets[pos].api)
        print(f"→ Лист перемещён на позицию {pos}")
    except Exception as e:
        print(f"⚠️ Не удалось переместить лист: {e}")

    rev_ws.range(1, 1).value = rev_hdr
    if rows:
        rev_ws.range(2, 1).value = rows
        last_row = len(rows) + 1
        total_col = len(rev_hdr)
        sum_row = last_row + 1
        rev_ws.range((sum_row, 1)).value = 'Итого'

        # Формулы для итогов по каждому месяцу и "Всего"
        for c in range(4, total_col+1):
            col_letter = xw.utils.col_name(c)
            rev_ws.range((sum_row, c)).formula = f'=SUM({col_letter}2:{col_letter}{last_row})'
            rev_ws.range((sum_row, c)).number_format = '#,##0 ₽'

        # Оформляем как умную таблицу
        table_range = rev_ws.range((1,1), (sum_row, total_col))
        format_table(rev_ws, table_range, TABLE_NAME)
        rev_ws.api.Application.ActiveWindow.SplitRow = 1
        rev_ws.api.Application.ActiveWindow.FreezePanes = True
        print('→ Данные и форматирование записаны')
    else:
        print('Нет данных для вывода — таблица не создаётся')

    if app: wb.save(); app.quit()
    print('=== Скрипт успешно завершён ===')

if __name__ == '__main__':
    main() 
