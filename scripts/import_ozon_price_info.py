# import_ozon_price_info.py

import os
import requests
import xlwings as xw
import pandas as pd
from time import sleep
from pathlib import Path
from scripts.sheet_utils import apply_sheet_settings


EXCEL_PATH = str(Path(__file__).resolve().parents[1] / 'Finmodel.xlsm')

SHEET_SETTINGS = 'НастройкиОрганизаций'
SHEET_PRICES   = 'ЦеныОзон'
API_URL        = 'https://api-seller.ozon.ru/v5/product/info/prices'

OUTPUT_HEADERS = [
    'Артикул','ID товара','Эквайринг max',
    'FBO: доставка','FBO: магистраль от','FBO: магистраль до','FBO: возвраты',
    'FBS: доставка','FBS: магистраль от','FBS: магистраль до',
    'FBS: первый километр мин','FBS: первый километр макс','FBS: возвраты',
    'FBO: % продажи','FBS: % продажи',
    'Валюта','Авто-акции включены','Авто-добавление в акции',
    'Цена с акциями','Цена продавца с акциями','Минимальная цена','Старая цена','Итоговая цена','Цена поставщика','НДС',
    'Акции есть','Акции период с','Акции период по'
]

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

def get_ozon_credentials(settings_ws):
    df = settings_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value

    print(f'=== Заголовки таблицы: {list(df.columns)}')
    print('=== Первые строки:')
    print(df.head(3).to_string(index=False))

    df.columns = [str(c).strip() for c in df.columns]
    print(f'=== После trim колонок: {list(df.columns)}')

    for col in ['Client-Id', 'Token_Ozon']:
        if col not in df.columns:
            raise Exception(f'❌ Отсутствует колонка "{col}"')

    found = df[df['Client-Id'].notna() & (df['Client-Id'] != '') & df['Token_Ozon'].notna() & (df['Token_Ozon'] != '')]
    print(f'=== Найдено строк с заполненными Client-Id и Token_Ozon: {len(found)}')
    if found.empty:
        raise Exception('❌ Не найдено ни одной строки с Client-Id и Token_Ozon')
    return found[['Client-Id', 'Token_Ozon']].dropna().values.tolist()

    print('=== Строка для API:')
    print(row)

    # --- КОРРЕКТНО ПРИВОДИМ К СТРОКЕ ---
    client_id = str(row['Client-Id']).strip()
    # если Client-Id float и выглядит как 142768.0 -> привести к 142768
    if client_id.endswith('.0'):
        client_id = client_id[:-2]
    api_key   = str(row['Token_Ozon']).strip()

    print(f"=== Итог: client_id='{client_id}', api_key (first 10)='{api_key[:10]}' ...")
    if not client_id or not api_key:
        raise Exception('❌ Не заданы Client-Id или Token_Ozon')
    return client_id, api_key

def main():
    print("=== Старт import_ozon_price_info ===")
    wb, app = get_workbook()
    try:
        settings_ws = wb.sheets[SHEET_SETTINGS]
        prices_ws = wb.sheets[SHEET_PRICES] if SHEET_PRICES in [s.name for s in wb.sheets] else wb.sheets.add(SHEET_PRICES)
        prices_ws.clear()
        prices_ws.range(1, 1).value = OUTPUT_HEADERS
        print(f'→ Лист {SHEET_PRICES} очищен и подготовлен')

        credentials = get_ozon_credentials(settings_ws)
        row_ptr = 2  # строка для записи (после заголовка)

        for org_idx, (client_id_raw, api_key) in enumerate(credentials, start=1):
            client_id = str(int(float(client_id_raw))).strip()
            print(f"\n=== Организация #{org_idx}: client_id='{client_id}' ===")

            session = requests.Session()
            session.headers.update({
                'Client-Id': client_id,
                'Api-Key': str(api_key).strip(),
                'Content-Type': 'application/json'
            })

            for vis in ['ALL', 'ARCHIVED']:
                print(f'→ Получение данных для видимости: {vis}')
                cursor = ''
                page = 1
                while True:
                    payload = {"filter": {"visibility": vis}, "limit": 1000}
                    if cursor:
                        payload["cursor"] = cursor
                    try:
                        resp = session.post(API_URL, json=payload, timeout=30)
                        print(f"  → page {page}, HTTP {resp.status_code}")
                        if resp.status_code != 200:
                            print(f'❌ Ошибка {resp.status_code}: {resp.text}')
                            break
                        result = resp.json()
                    except Exception as e:
                        print(f'❌ Ошибка при запросе: {e}')
                        break

                    cursor = result.get('cursor', '')
                    items = result.get('items', [])
                    print(f'    Строк в ответе: {len(items)}')
                    if not items:
                        break

                    rows = []
                    for it in items:
                        c = it.get('commissions', {}) or {}
                        p = it.get('price', {}) or {}
                        m = it.get('marketing_actions', {}) or {}
                        rows.append([
                            it.get('offer_id', ''), it.get('product_id', ''), c.get('acquiring'),
                            c.get('fbo_deliv_to_customer_amount'), c.get('fbo_direct_flow_trans_min_amount'), c.get('fbo_direct_flow_trans_max_amount'), c.get('fbo_return_flow_amount'),
                            c.get('fbs_deliv_to_customer_amount'), c.get('fbs_direct_flow_trans_min_amount'), c.get('fbs_direct_flow_trans_max_amount'),
                            c.get('fbs_first_mile_min_amount'), c.get('fbs_first_mile_max_amount'), c.get('fbs_return_flow_amount'),
                            c.get('sales_percent_fbo'), c.get('sales_percent_fbs'),
                            p.get('currency_code'), p.get('auto_action_enabled'), p.get('auto_add_to_ozon_actions_list_enabled'),
                            p.get('marketing_price'), p.get('marketing_seller_price'), p.get('min_price'), p.get('old_price'), p.get('price'), p.get('retail_price'), p.get('vat'),
                            m.get('ozon_actions_exist'), m.get('current_period_from'), m.get('current_period_to')
                        ])
                    if rows:
                        prices_ws.range(row_ptr, 1).value = rows
                        row_ptr += len(rows)
                    page += 1
                    if not cursor:
                        break
                sleep(0.5)

        # Итоги
        prices_ws.range('A1').expand().columns.autofit()
        prices_ws.api.Rows(1).Font.Bold = True
        prices_ws.activate()
        prices_ws.range('A2').select()
        wb.app.api.ActiveWindow.FreezePanes = True

        # Удалить предыдущие таблицы
        for tbl in prices_ws.tables:
            if tbl.name == "OzonPricesTable":
                tbl.delete()

        last_row = prices_ws.range('A1').end('down').row
        last_col = len(OUTPUT_HEADERS)
        tbl_range = prices_ws.range((1, 1), (last_row, last_col))
        prices_ws.tables.add(tbl_range, name="OzonPricesTable", table_style_name="TableStyleMedium7", has_headers=True)
        print("→ Умная таблица создана (TableStyleMedium7)")

        apply_sheet_settings(wb, SHEET_PRICES)

        print(f'→ Итог: записано строк: {row_ptr-2}')
    except Exception as e:
        print(f'❌ Ошибка: {e}')
    finally:
        if app:
            wb.save()
            app.quit()
        print('=== Скрипт успешно завершён ===')

if __name__ == '__main__':
    main()
