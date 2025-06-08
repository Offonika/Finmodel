# update_plan_sales_ozon.py

import os
import xlwings as xw
import pandas as pd
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')

SHEET_SETTINGS = 'НастройкиОрганизаций'
SHEET_SEASON   = 'Сезонность'
SHEET_SALES    = 'ФинотчетыОзон'
SHEET_PRICES   = 'ЦеныОзон'
SHEET_PLAN     = 'ПланПродажОзон'
TABLE_NAME     = 'PlanOzonTable'
TABLE_STYLE    = 'TableStyleMedium7'

MONTHS_CNT = 12
MONTH_NAMES = [f'Мес.{str(i+1).zfill(2)}' for i in range(MONTHS_CNT)]
CURRENT_MONTH = datetime.now().month

def normalize_sku(val):
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def col_to_letter(col):
    letter = ''
    while col > 0:
        col, rem = divmod(col-1, 26)
        letter = chr(65 + rem) + letter
    return letter

def safe_float(val):
    if pd.isna(val):
        return 0.0
    try:
        return float(str(val).replace(',', '.').replace(' ','').replace(' ','')) # Убирает пробелы и неразрывные
    except Exception:
        return 0.0

def main():
    print("=== Старт update_plan_sales_ozon ===")
    wb, app = get_workbook()
    print('→ Открыт файл:', wb.fullname)

    # 1. Сезонные коэффициенты по SKU
    try:
        season_ws = wb.sheets[SHEET_SEASON]
        season_df = season_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
        print(f'→ Лист {SHEET_SEASON} считан: {len(season_df)} строк')
    except Exception as e:
        print(f'❌ Нет листа {SHEET_SEASON}')
        if app: app.quit()
        return
    season_factors = {}
    for _, r in season_df.iterrows():
        sku = str(r.iloc[0])
        vals = [safe_float(r.iloc[i]) if i < len(r) else 1.0 for i in range(1, MONTHS_CNT+1)]
        season_factors[sku] = vals

    # 2. История продаж (ФинотчетыОзон)
    try:
        sales_ws = wb.sheets[SHEET_SALES]
        sales_df = sales_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
        print(f'→ Лист {SHEET_SALES} считан: {len(sales_df)} строк')
    except Exception as e:
        print(f'❌ Нет листа {SHEET_SALES}')
        if app: app.quit()
        return
    idx_sales = {h: i for i, h in enumerate(sales_df.columns)}
    req_cols = ['Организация','Артикул_поставщика','SKU','Продано шт.','Месяц']
    for col in req_cols:
        if col not in idx_sales:
            print(f'❌ Нет колонки "{col}" в {SHEET_SALES}')
            if app: app.quit()
            return

    CURRENT_YEAR = datetime.now().year  # или любой нужный

    sku_to_offer = {}
    for _, r in sales_df.iterrows():
        org = str(r['Организация'])
        sku = normalize_sku(r['SKU'])
        offer = str(r['Артикул_поставщика'])
        if sku and offer:
            sku_to_offer[(org, sku)] = offer

    qty_map = {}
    for _, r in sales_df.iterrows():
        year = int(safe_float(r['Год']))
        if year != CURRENT_YEAR:
            continue
        org  = str(r['Организация'])
        sku = normalize_sku(r['SKU'])
        month = int(safe_float(r['Месяц']))
        qty = safe_float(r['Продано шт.'])
        key = (org, sku)
        if key not in qty_map:
            qty_map[key] = [0.0] * MONTHS_CNT
        if 1 <= month <= MONTHS_CNT:
            qty_map[key][month-1] += qty



    # 3. Цены по "ЦеныОзон"
    try:
        prices_ws = wb.sheets[SHEET_PRICES]
        prices_df = prices_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
        print(f'→ Лист {SHEET_PRICES} считан: {len(prices_df)} строк')
    except Exception as e:
        print(f'❌ Нет листа {SHEET_PRICES}')
        if app: app.quit()
        return
    idx_price = {h: i for i, h in enumerate(prices_df.columns)}
    if 'Артикул' not in idx_price or 'Цена продавца с акциями' not in idx_price:
        print(f'❌ Нет нужных колонок в {SHEET_PRICES}')
        if app: app.quit()
        return
    price_map = {}
    for _, r in prices_df.iterrows():
        offer = str(r['Артикул'])
        price_map[offer] = safe_float(r['Цена продавца с акциями'])

   
    # 4. Сборка плана
    rows = []
    for key, hist in qty_map.items():
        org, sku = key
        offer = sku_to_offer.get((org, sku), '')   # находим offer для (org, sku)
        total_hist = sum(hist[:CURRENT_MONTH])
        if total_hist == 0:
            continue
        base = round(total_hist / max(1, CURRENT_MONTH))
        factors = season_factors.get(sku, [1.0] * MONTHS_CNT)
        price = price_map.get(offer, 0.0)          # цены по offer
        plan = []
        for i in range(MONTHS_CNT):
            if i < CURRENT_MONTH-1:
                plan.append(round(hist[i]))
            else:
                plan.append(round(base * factors[i]))
        if sum(plan) == 0:
            continue
        row = [org, offer, sku, base, price] + plan
        rows.append(row)

    # Сортировка по сумме продаж
    rows.sort(key=lambda r: -sum(r[5:5+MONTHS_CNT]))

    # 5. Вывод на лист
    try:
        plan_ws = wb.sheets[SHEET_PLAN]
        plan_ws.clear()
        print(f'→ Лист {SHEET_PLAN} очищен')
    except:
        plan_ws = wb.sheets.add(SHEET_PLAN)
        print(f'→ Лист {SHEET_PLAN} создан')

    header = ['Организация','Артикул_поставщика','SKU','Базовое кол-во','Плановая цена'] + MONTH_NAMES + ['Всего']
    plan_ws.range(1,1).value = header

    # ----- Установка цвета ярлыка и позиция листа -----
    try:
        plan_ws.api.Tab.Color = (0, 192, 255)  # BGR!
        if plan_ws.index != 3:
            plan_ws.api.Move(Before=wb.sheets[2].api)
        print("→ Установлен цвет ярлыка #FFC000 и позиция №3")
    except Exception as e:
        print(f"⚠️ Не удалось установить цвет/позицию листа: {e}")

    # Вставляем строки (в "Всего" формула)
    values = []
    for i, r in enumerate(rows):
        row_num = i + 2
        col_start = header.index('Мес.01') + 1
        col_end = header.index('Мес.12') + 1
        col_letter_start = col_to_letter(col_start)
        col_letter_end = col_to_letter(col_end)
        sum_formula = f'=SUM({col_letter_start}{row_num}:{col_letter_end}{row_num})'
        values.append(r + [sum_formula])
    if values:
        plan_ws.range(2, 1).value = values

    # Итоговая строка "Итого"
    last_row = len(values) + 2
    total_row = []
    for j in range(len(header)):
        if j < 5:
            total_row.append('Итого' if j == 0 else '')
        else:
            col_letter = col_to_letter(j+1)
            total_row.append(f'=SUM({col_letter}2:{col_letter}{last_row-1})')
    plan_ws.range(last_row, 1).value = total_row

    # Форматирование как умная таблица
    for tbl in plan_ws.tables:
        if tbl.name == TABLE_NAME:
            tbl.delete()
    table_range = plan_ws.range((1, 1), (last_row, len(header)))
    plan_ws.tables.add(table_range, name=TABLE_NAME, table_style_name=TABLE_STYLE, has_headers=True)
    plan_ws.range('A1').expand().columns.autofit()
    plan_ws.api.Rows(1).Font.Bold = True
    plan_ws.api.Application.ActiveWindow.SplitRow = 1
    plan_ws.api.Application.ActiveWindow.FreezePanes = True

    print('=== Скрипт успешно завершён ===')
    if app: wb.save(); app.quit()


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

if __name__ == '__main__':
    main()
