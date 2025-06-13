# import_ozon_realization_grouped.py

import warnings
from collections import defaultdict

warnings.filterwarnings("ignore", category=UserWarning)
import requests
import pandas as pd
import glob
import os            # ← импорт был здесь
import xlwings as xw
from datetime import datetime

def get_workbook():
    try:
        wb = xw.Book.caller()   # Если вызов из Excel
        return wb, wb.app, False
    except Exception:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(EXCEL_PATH)
        return wb, app, True



# === Константы ===
OUTPUT_HEADERS = [
    'Организация','Год','Месяц','Артикул_поставщика','SKU','Штрихкод','Название товара',
    'Сумма продаж ед.','Сумма коэфф. комиссий','Дост: сумма','Дост: бонус',
    'Дост: комиссия','Дост: компенсация','Дост: цена/ед.','Дост: кол-во',
    'Дост: стандарт. вознагражд.','Дост: соинвест. банка','Дост: звёзды',
    'Дост: соинвест. ПВЗ','Дост: итого','Возв: сумма','Возв: бонус',
    'Возв: комиссия','Возв: компенсация','Возв: цена/ед.','Возв: кол-во',
    'Возв: стандарт. вознагражд.','Возв: соинвест. банка','Возв: звёзды',
    'Возв: соинвест. ПВЗ','Возв: итого','Продано шт.','Реализовано (руб)',
    'Всего выплат от партнёров','Начислено баллов','Базовое вознаграждение Ozon',
    'Вознаграждение после скидок'
]
API_URL = 'https://api-seller.ozon.ru/v2/finance/realization/'
SHEET_NAME = 'ФинотчетыОзон'
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')
  # подстрой под себя!
ORG_SHEET = 'НастройкиОрганизаций'
CFG_SHEET = 'Настройки'

def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

def load_settings(cfg_path):
    # 1. Читаем верхнюю часть листа организаций до первой пустой строки
    df_all = pd.read_excel(cfg_path, sheet_name=ORG_SHEET, header=0)
    org_rows = []
    for i, row in df_all.iterrows():
        if pd.isna(row['Организация']) or str(row['Организация']).strip() == '':
            break
        org_rows.append(row)
    if not org_rows:
        raise Exception("В таблице 'НастройкиОрганизаций' нет данных об организациях!")
    df_orgs = pd.DataFrame(org_rows)

    def safe_str(x):
        if pd.isna(x):
            return ''
        s = str(x).strip()
        if s.endswith('.0'):
            s = s[:-2]
        return s

    orgs = []
    for _, row in df_orgs.iterrows():
        org = safe_str(row['Организация'])
        client_id = safe_str(row['Client-Id'])
        token = safe_str(row['Token_Ozon'])
        print(f"[DEBUG] Организация: {org} | Client-Id: {client_id} | Api-Key: {token}")  # отладка!
        if org and client_id and token:
            orgs.append({'org': org, 'client_id': client_id, 'token': token})

    # 2. Считываем периоды загрузки из листа "Настройки"
    df_params = pd.read_excel(cfg_path, sheet_name=CFG_SHEET, header=None)
    period_start, period_end = None, None
    for _, row in df_params.iterrows():
        if str(row[0]).strip() == 'ПериодНачало':
            period_start = pd.to_datetime(row[1])
        if str(row[0]).strip() == 'ПериодКонец':
            period_end = pd.to_datetime(row[1])
    if period_start is None:
        raise Exception("Не задана дата ПериодНачало в таблице параметров!")
    if period_end is None:
        period_end = pd.Timestamp.today()
    return orgs, period_start, period_end

def get_periods(start, end):
    periods = []
    cur = pd.Timestamp(start.year, start.month, 1)
    last = pd.Timestamp(end.year, end.month, 1)
    while cur <= last:
        periods.append({'year': cur.year, 'month': cur.month})
        cur += pd.DateOffset(months=1)
    return periods

def fetch_ozon_data(org, client_id, token, year, month):
    headers = {'Client-Id': client_id, 'Api-Key': token}
    body = {'year': year, 'month': month}
    resp = requests.post(API_URL, headers=headers, json=body)
    if resp.status_code != 200:
        log(f"Ошибка API {org} {year}-{month}: {resp.status_code}")
        return []
    return resp.json().get('result', {}).get('rows', [])

def format_as_table(ws, df, table_name="OzonReportTable"):
    last_row = df.shape[0] + 1  # +1 для шапки
    last_col = df.shape[1]
    rng = ws.range((1, 1), (last_row, last_col))
    # Удаляем предыдущую таблицу если была
    for tbl in ws.tables:
        if tbl.name == table_name:
            tbl.delete()
    # Преобразуем в таблицу Excel (умную таблицу)
    ws.tables.add(
        rng,
        name=table_name,
        table_style_name="TableStyleMedium7",  # Зелёный, Medium 7
        has_headers=True
    )

    # Жирная шапка и автоширина
    ws.range((1, 1), (1, last_col)).api.Font.Bold = True
    ws.range((1, 1), (last_row, last_col)).columns.autofit()
    # Границы
    ws.range((1, 1), (last_row, last_col)).api.Borders.Weight = 2  # xlThin
    # Фиксируем первую строку (без Select)
    ws.api.Activate()  # Активируем лист (иначе FreezePanes не работает)
    ws.api.Application.ActiveWindow.SplitRow = 1
    ws.api.Application.ActiveWindow.FreezePanes = True


def write_to_excel(df: pd.DataFrame):
    wb, app, created = get_workbook()
    try:
        # Лист очищаем или добавляем
        if SHEET_NAME in [s.name for s in wb.sheets]:
            ws = wb.sheets[SHEET_NAME]
            ws.clear_contents()
        else:
            ws = wb.sheets.add(SHEET_NAME)

 # Установка цвета ярлыка #84F8EA (бирюзовый)
        try:
            ws.api.Tab.Color = 0xEAF884  # BGR для xlwings/COM
            print("→ Цвет ярлыка #84F8EA установлен")
        except Exception as e:
            print(f"⚠️ Не удалось установить цвет ярлыка: {e}")

        # Переместить лист на позицию 11
        try:
            if ws.index != 11:
                ws.api.Move(Before=ws.book.sheets[17].api)
                print("→ Лист перемещён на позицию 17")
        except Exception as e:
            print(f"⚠️ Не удалось переместить лист: {e}")


        ws.range('A1').options(index=False, header=True).value = df
        format_as_table(ws, df)
        wb.save()
    finally:
        if created:   # Только если мы сами открывали Excel — закрываем!
            wb.close()
            app.quit()

def main():
    log(f"Старт скрипта. Excel: {EXCEL_PATH}")
    orgs, period_start, period_end = load_settings(EXCEL_PATH)
    periods = get_periods(period_start, period_end)
    log(f"Организаций для запроса: {len(orgs)}")
    log(f"Периоды: {period_start:%Y-%m} — {period_end:%Y-%m}")

    groups = defaultdict(lambda: {
        'org': '', 'year': 0, 'month': 0,
        'offer_id': '', 'sku': '', 'barcode': '', 'name': '',
        'commission_ratio': 0, 'seller_price_sum': 0,
        'del': defaultdict(float), 'ret': defaultdict(float)
    })
    for org_info in orgs:
        org = org_info['org']
        client_id = org_info['client_id']
        token = org_info['token']
        for p in periods:
            log(f"Запрос: {org} {p['year']}-{p['month']}")
            rows = fetch_ozon_data(org, client_id, token, p['year'], p['month'])
            for r in rows:
                it = r.get('item', {})
                deliv = r.get('delivery_commission') or {}
                ret = r.get('return_commission') or {}
                key = (org, p['year'], p['month'], it.get('offer_id', ''), it.get('sku', ''), it.get('barcode', ''), it.get('name', ''))
                g = groups[key]
                g['org'], g['year'], g['month'], g['offer_id'], g['sku'], g['barcode'], g['name'] = key
                g['commission_ratio'] += r.get('commission_ratio', 0)
                g['seller_price_sum'] += r.get('seller_price_per_instance', 0)
                for src, dct in zip(['del', 'ret'], [deliv, ret]):
                    for f_api, f_out in [
                        ('amount', 'amount'), ('bonus', 'bonus'), ('commission', 'commission'), ('compensation', 'compensation'),
                        ('price_per_instance', 'price'), ('quantity', 'qty'), ('standard_fee', 'std_fee'),
                        ('bank_coinvestment', 'bank'), ('stars', 'stars'), ('pick_up_point_coinvestment', 'pvz'), ('total', 'total')
                    ]:
                        g[src][f_out] += dct.get(f_api, 0) or 0

    # --- Подготовка итоговой таблицы
    result = []
    for g in groups.values():
        sold_qty         = g['del']['qty'] - g['ret']['qty']
        partner_payouts  = g['del']['bank'] + g['del']['pvz'] - g['ret']['bank'] - g['ret']['stars']
        revenue          = g['del']['amount'] - g['ret']['amount'] + partner_payouts
        bonused_points   = g['del']['bonus'] - g['ret']['bonus']
        base_reward      = g['del']['std_fee'] - g['ret']['std_fee']
        net_reward       = base_reward - bonused_points
        result.append([
            g['org'], g['year'], g['month'], g['offer_id'], g['sku'], g['barcode'], g['name'],
            g['seller_price_sum'], g['commission_ratio'],
            g['del']['amount'], g['del']['bonus'], g['del']['commission'], g['del']['compensation'], g['del']['price'], g['del']['qty'], g['del']['std_fee'], g['del']['bank'], g['del']['stars'], g['del']['pvz'], g['del']['total'],
            g['ret']['amount'], g['ret']['bonus'], g['ret']['commission'], g['ret']['compensation'], g['ret']['price'], g['ret']['qty'], g['ret']['std_fee'], g['ret']['bank'], g['ret']['stars'], g['ret']['pvz'], g['ret']['total'],
            sold_qty, revenue, partner_payouts, bonused_points, base_reward, net_reward
        ])
    df = pd.DataFrame(result, columns=OUTPUT_HEADERS)

    write_to_excel(df)   # ← Вот так

if __name__ == '__main__':
    main()
