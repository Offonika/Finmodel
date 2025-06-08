# calculate_cogs_batched.py

import os
import xlwings as xw
import pandas as pd
import math

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')

# ЗАМЕНЁННЫЕ НАЗВАНИЯ ЛИСТОВ
SHEET_PRODUCTS = 'Номенклатура_WB'
SHEET_PRICES   = 'ЗакупочныеЦены'
SHEET_DUTIES   = 'ТаможенныеПошлины'
SHEET_SETTINGS = 'Настройки'
SHEET_RESULT   = 'РасчётСебестоимости'
TABLE_NAME     = 'CogsTable'
TABLE_STYLE    = 'TableStyleMedium7'
PROGRESS_CELL  = 'Z1'  # Можно скрыто хранить прогресс для батча

BATCH_SIZE = 1000  # Объём одной порции для записи в Excel

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
        if pd.isna(val): return 0.0
        return float(str(val).replace(',', '.').replace(' ', '').replace(' ', ''))
    except Exception:
        return 0.0

def read_settings(ws):
    df = ws.range(1, 1).expand().options(pd.DataFrame, header=1, index=False).value
    df = df.loc[:, ~df.columns.duplicated()]  # Убираем дубликаты
    idx = {h: i for i, h in enumerate(df.columns)}
    vals = df.values.tolist()
    params = {}
    for row in vals:
        param = str(row[0])
        val = row[1] if len(row) > 1 else None
        if not param: break
        params[param] = val

    def get_num(name, default=0):
        v = params.get(name, default)
        if v is None: return default
        try:
            return float(str(v).replace(',', '.').replace('%','').replace(' ',''))
        except:
            return default

    return {
        "cargoRatePerKg": get_num('Логистика_Карго_$/кг'),
        "whiteRatePerKg": get_num('Логистика_Белая_$/кг'),
        "usdRate": get_num('Курс_USD'),
        "cnyRate": get_num('Курс_CNY'),
        "ndsRateWhite": get_num('НДС_Белая', 0) / 100.0 if get_num('НДС_Белая', 0) > 1 else get_num('НДС_Белая', 0)
    }

def get_logistics_mode(org, ws):
    vals = ws.range(1, 1).expand().options(pd.DataFrame, header=1, index=False).value
    row = vals[vals.iloc[:, 0] == org]
    if not row.empty and 'Тип_Логистики' in row.columns:
        val = row.iloc[0]['Тип_Логистики']
        if isinstance(val, str) and 'бел' in val.lower():
            return 'Белая'
    return 'Карго'

def get_progress(ws):
    try:
        val = ws.range(PROGRESS_CELL).value
        return int(val) if val else 1
    except:
        return 1

def set_progress(ws, idx):
    ws.range(PROGRESS_CELL).value = idx

def clear_progress(ws):
    ws.range(PROGRESS_CELL).value = None

def main():
    print('=== Старт batch расчёта себестоимости ===')
    wb, app = get_workbook()

    try:
        prod_ws = wb.sheets[SHEET_PRODUCTS]
        price_ws = wb.sheets[SHEET_PRICES]
        duty_ws = wb.sheets[SHEET_DUTIES]
        settings_ws = wb.sheets[SHEET_SETTINGS]
    except Exception as e:
        print(f'❌ Не найден один из листов: {e}')
        if app: app.quit()
        return

    global_params = read_settings(settings_ws)
    print(f'→ Параметры: {global_params}')

    prod_df = prod_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
    price_df = price_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
    duty_df = duty_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value

    idxP = {h: i for i, h in enumerate(prod_df.columns)}
    idxC = {h: i for i, h in enumerate(price_df.columns)}
    idxD = {h: i for i, h in enumerate(duty_df.columns)}

    price_dict = {str(r['Артикул_поставщика']).strip().upper(): r for _, r in price_df.iterrows()}

    duty_dict = {str(r['Предмет']): r for _, r in duty_df.iterrows()}

    result_ws = None
    try:
        result_ws = wb.sheets[SHEET_RESULT]
        header_row = result_ws.range(1,1).expand('right').value
        print(f'→ Лист {SHEET_RESULT} найден, строк: {result_ws.range("A1").end("down").row}')
    except:
        result_ws = wb.sheets.add(SHEET_RESULT)
        header_row = None
        print(f'→ Лист {SHEET_RESULT} создан')

    header = [
        'Организация',
        'Артикул_поставщика',
        'Предмет',
        'Наименование',
        'Закуп_Цена_руб',
        'Логистика_руб',
        'Пошлина_руб',
        'НДС_руб',
        'Себестоимость_руб',
        'Себестоимость_без_НДС_руб',
        'Входящий_НДС_руб'
    ]
    start_idx = get_progress(result_ws)
    if start_idx == 1 or not header_row or header_row != header:
        result_ws.clear()
        result_ws.range(1,1).value = header
        start_idx = 1
        print('→ Заголовок записан, таблица очищена')
        set_progress(result_ws, 1)

    # --- Диагностика ---
    print(f'Всего товаров: {len(prod_df)}, Закупочных цен: {len(price_df)}')

    # Список артикулов из price_dict
    all_price_keys = set(price_dict.keys())

    # Список артикулов из Номенклатуры
    all_product_keys = set(str(x).strip().upper() for x in prod_df['Артикул_поставщика'])

    # Артикулы без закупочной цены
    not_found = all_product_keys - all_price_keys
    print(f'❗ Не найдены закупочные цены для {len(not_found)} товаров. Примеры: {list(not_found)[:10]}')



    n = len(prod_df)
    print(f'→ Всего товаров: {n}, стартовый индекс: {start_idx}')

    processed = 0
    skipped = 0
    idx = start_idx - 1
    while idx < n:
        batch_end = min(n, idx + BATCH_SIZE)
        batch = []
        for i in range(idx, batch_end):
            r = prod_df.iloc[i]
            org = r['Организация']
            vendor = str(r['Артикул_поставщика']).strip().upper()

            subject = r['Предмет']
            name = r['Название']
            weight = safe_float(r['Вес_брутто'])

            price_row = price_dict.get(str(vendor))
            if price_row is None or (hasattr(price_row, 'empty') and price_row.empty):
                print(f'Skip {vendor} ({org}) – no purchase price found')
                skipped += 1
                continue

            price_val = safe_float(price_row.get('Закуп_Цена'))
            currency = price_row.get('Валюта')
            if currency == 'USD':
                rate = global_params['usdRate']
            elif currency == 'CNY':
                rate = global_params['cnyRate']
            else:
                rate = 1
            purchase_rub = price_val * rate

            duty_row = duty_dict.get(subject)
            logistics_mode = get_logistics_mode(org, settings_ws)
            logistics_rate_per_kg = global_params['cargoRatePerKg'] if logistics_mode == 'Карго' else global_params['whiteRatePerKg']
            logistics_rub = weight * logistics_rate_per_kg * global_params['usdRate']

            duty_rate = 0
            if logistics_mode == 'Белая' and duty_row is not None:
                duty_rate_val = duty_row.get('Пошлина')
                if isinstance(duty_rate_val, str):
                    duty_rate = float(duty_rate_val.replace('%', '').replace(',', '.')) / 100
                elif isinstance(duty_rate_val, (int, float)):
                    duty_rate = duty_rate_val / 100 if duty_rate_val > 1 else duty_rate_val
            duty_rub = purchase_rub * duty_rate

            vat_rub = (purchase_rub + duty_rub + logistics_rub) * global_params['ndsRateWhite'] if logistics_mode == 'Белая' else 0
            total_cogs = purchase_rub + duty_rub + logistics_rub + vat_rub
            cogs_without_vat = total_cogs - vat_rub

            batch.append([
                org,
                vendor,
                subject,
                name,
                round(purchase_rub),
                round(logistics_rub),
                round(duty_rub),
                round(vat_rub),
                round(total_cogs),
                round(cogs_without_vat),
                round(vat_rub)
            ])
            processed += 1

        if batch:
            first_row = result_ws.range('A1').end('down').row + 1 if result_ws.range('A1').end('down').row > 1 else 2
            result_ws.range(first_row, 1).value = batch
            print(f'→ В таблицу добавлено строк: {len(batch)}')
        else:
            print('→ Нет новых строк для записи')

        idx = batch_end
        set_progress(result_ws, idx + 1)

    clear_progress(result_ws)
    for tbl in result_ws.tables:
        if tbl.name == TABLE_NAME:
            tbl.delete()
    last_row = result_ws.range('A1').end('down').row
    rng = result_ws.range((1,1), (last_row, len(header)))
    result_ws.tables.add(rng, name=TABLE_NAME, table_style_name=TABLE_STYLE, has_headers=True)
    result_ws.range('A1').expand().columns.autofit()
    result_ws.api.Rows(1).Font.Bold = True
    print(f'→ Расчёт завершён и таблица стилизована (итого строк: {last_row-1})')
    if app:
        wb.save(); app.quit()
    print('=== Скрипт завершён ===')

if __name__ == '__main__':
    main()
