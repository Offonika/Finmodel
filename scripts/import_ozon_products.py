# import_ozon_products.py
"""Load product list from Ozon and merge into sheet "Номенклатура_WB".

The workbook must contain sheet "НастройкиОрганизаций" with columns
"Организация", "Client-Id" and "Token_Ozon". The sheet
"Номенклатура_WB" uses the same columns as the Wildberries loader and
is updated/extended with Ozon products.
"""

import os
import requests
import pandas as pd
import xlwings as xw

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')

SETTINGS_SHEET = 'НастройкиОрганизаций'
PRODUCTS_SHEET = 'Номенклатура_WB'
HEADERS = [
    'Организация', 'Артикул_WB', 'Артикул_поставщика',
    'Бренд', 'Название', 'Предмет',
    'Ширина', 'Высота', 'Длина', 'Вес_брутто', 'Объем_литр'
]
API_URL = 'https://api-seller.ozon.ru/v3/product/list'
PAGE_LIMIT = 1000


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


def read_credentials(ws):
    df = ws.range(1, 1).expand().options(pd.DataFrame, header=1, index=False).value
    df.columns = [str(c).strip() for c in df.columns]
    creds = []
    for _, row in df.iterrows():
        org = str(row.get('Организация', '')).strip()
        cid = str(row.get('Client-Id', '')).strip()
        api_key = str(row.get('Token_Ozon', '')).strip()
        if cid.endswith('.0'):
            cid = cid[:-2]
        if org and cid and api_key:
            creds.append({'org': org, 'client_id': cid, 'api_key': api_key})
    return creds


def fetch_products(session):
    last_id = ''
    page = 1
    items = []
    while True:
        payload = {'limit': PAGE_LIMIT, 'filter': {'visibility': 'ALL'}}
        if last_id:
            payload['last_id'] = last_id
        try:
            resp = session.post(API_URL, json=payload, timeout=30)
            print(f'  → page {page}, HTTP {resp.status_code}')
            if resp.status_code != 200:
                print(f'    ❌ Ошибка {resp.status_code}: {resp.text}')
                break
            data = resp.json().get('result', {})
        except Exception as e:
            print(f'    ❌ Ошибка запроса: {e}')
            break
        batch = data.get('items') or data.get('products') or []
        print(f'    Строк в ответе: {len(batch)}')
        items.extend(batch)
        last_id = data.get('last_id')
        if not last_id or not batch:
            break
        page += 1
    return items


def merge_products(df_old: pd.DataFrame, df_new: pd.DataFrame) -> pd.DataFrame:
    key = ['Организация', 'Артикул_поставщика']
    if df_old.empty:
        return df_new
    merged = pd.merge(df_old, df_new, on=key, how='outer', suffixes=('', '_new'))
    for col in ['Артикул_WB', 'Название']:
        if f'{col}_new' in merged.columns:
            merged[col] = merged[f'{col}_new'].combine_first(merged[col])
            merged.drop(columns=[f'{col}_new'], inplace=True)
    return merged[HEADERS]


def main():
    print('=== Старт import_ozon_products ===')
    wb, app = get_workbook()
    try:
        settings_ws = wb.sheets[SETTINGS_SHEET]
    except Exception:
        print(f'❌ Нет листа {SETTINGS_SHEET}')
        if app:
            app.quit()
        return
    try:
        prod_ws = wb.sheets[PRODUCTS_SHEET]
    except Exception:
        prod_ws = wb.sheets.add(PRODUCTS_SHEET)
        prod_ws.range(1, 1).value = HEADERS
    if prod_ws.range('A1').value != HEADERS[0]:
        prod_ws.clear()
        prod_ws.range(1, 1).value = HEADERS

    df_existing = prod_ws.range(1, 1).expand().options(pd.DataFrame, header=1, index=False).value
    if df_existing is None or df_existing.empty:
        df_existing = pd.DataFrame(columns=HEADERS)

    creds = read_credentials(settings_ws)
    if not creds:
        print('❌ Нет организаций с Client-Id и Token_Ozon')
        if app:
            app.quit()
        return

    df_all_new = pd.DataFrame(columns=HEADERS)

    for idx, info in enumerate(creds, start=1):
        print(f"→ Организация {info['org']} ({idx}/{len(creds)})")
        session = requests.Session()
        session.headers.update({
            'Client-Id': info['client_id'],
            'Api-Key': info['api_key'],
            'Content-Type': 'application/json'
        })
        items = fetch_products(session)
        rows = []
        for it in items:
            offer = str(it.get('offer_id', '')).strip()
            prod_id = it.get('product_id', '')
            name = it.get('name', '')
            rows.append([
                info['org'], prod_id, offer,
                '', name, '', '', '', '', '', ''
            ])
        if rows:
            df_org = pd.DataFrame(rows, columns=HEADERS)
            df_all_new = pd.concat([df_all_new, df_org], ignore_index=True)
        session.close()

    df_result = merge_products(df_existing, df_all_new)

    prod_ws.clear()
    prod_ws.range(1, 1).value = HEADERS
    if not df_result.empty:
        prod_ws.range(2, 1).value = df_result.values
    prod_ws.range('A1').expand().columns.autofit()
    prod_ws.api.Rows(1).Font.Bold = True
    prod_ws.api.Application.ActiveWindow.SplitRow = 1
    prod_ws.api.Application.ActiveWindow.FreezePanes = True

    if app:
        wb.save()
        app.quit()
    print('=== Скрипт завершён ===')


if __name__ == '__main__':
    main()
