# calculate_redemption_rate.py
# ----------------------------------------------------------
# Расчёт % выкупа по nmId на основе srid за 90 дней
# ----------------------------------------------------------

import os
import time
import datetime
import requests
import pandas as pd
import xlwings as xw

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')
SHEET_SETTINGS = 'НастройкиОрганизаций'
SHEET_OUTPUT = '%ВыкупаWB'
DAYS = 90
WB_ORDERS_URL = 'https://statistics-api.wildberries.ru/api/v1/supplier/orders'
WB_SALES_URL = 'https://statistics-api.wildberries.ru/api/v1/supplier/sales'

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

def fetch_wb_data(url, token, date_from):
    headers = {'Authorization': token}
    data_all = []
    while True:
        response = requests.get(url, headers=headers, params={'dateFrom': date_from})
        if response.status_code != 200:
            print(f'⚠ Ошибка {response.status_code}: {response.text}')
            break
        data = response.json()
        if not data:
            break
        data_all.extend(data)
        date_from = data[-1]['lastChangeDate']
        print(f'  → Загружено: {len(data_all)} записей, продолжаем с {date_from}')
        time.sleep(60)
    return data_all

def process_org(name, token):
    print(f'\n=== Обработка: {name} ===')
    date_from = (datetime.datetime.now() - datetime.timedelta(days=DAYS)).strftime('%Y-%m-%dT00:00:00')
    
    print('→ Выгрузка заказов...')
    orders = fetch_wb_data(WB_ORDERS_URL, token, date_from)
    orders_filtered = {o['srid']: o['nmId'] for o in orders if not o.get('isCancel', True)}
    print(f'  → Отфильтровано заказов: {len(orders_filtered)}')

    print('→ Выгрузка продаж...')
    sales = fetch_wb_data(WB_SALES_URL, token, date_from)
    sales_srid = {s['srid'] for s in sales if s.get('saleID')}
    print(f'  → Продаж найдено: {len(sales_srid)}')

    result = {}
    for srid, nmId in orders_filtered.items():
        if nmId not in result:
            result[nmId] = {'orders': 0, 'sales': 0}
        result[nmId]['orders'] += 1
        if srid in sales_srid:
            result[nmId]['sales'] += 1

    print('→ Расчёт % выкупа...')
    rows = []
    for nmId, stats in result.items():
        percent = round(stats['sales'] / stats['orders'] * 100, 2) if stats['orders'] > 0 else 0
        rows.append([name, nmId, stats['orders'], stats['sales'], percent])

    return rows

def main():
    wb, app = get_workbook()
    sht = wb.sheets[SHEET_SETTINGS]
    df_settings = sht.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value

    all_results = []
    for _, row in df_settings.iterrows():
        token = row.get('Token_WB')
        name = row.get('Организация')
        if pd.isna(token) or pd.isna(name):
            continue
        try:
            rows = process_org(name, token)
            all_results.extend(rows)
        except Exception as e:
            print(f'❌ Ошибка при обработке {name}: {e}')

    df_result = pd.DataFrame(all_results, columns=[
        'Организация', 'nmId', 'Кол-во заказов', 'Кол-во продаж', '% выкупа'
    ])
    print('\n→ Запись в Excel...')
    
    if SHEET_OUTPUT in [s.name for s in wb.sheets]:
        sht_out = wb.sheets[SHEET_OUTPUT]
        sht_out.clear()
    else:
        sht_out = wb.sheets.add(SHEET_OUTPUT, after=wb.sheets[len(wb.sheets)-1])
    
    sht_out.range('A1').value = df_result
    sht_out.api.Tab.ColorIndex = 4  # зелёный ярлык
    sht_out.api.Move(Before=wb.sheets[32].api)  # вставить на 33 позицию

    print('✅ Готово.')
    if app is not None:
        wb.save()
        app.quit()

if __name__ == '__main__':
    main()
