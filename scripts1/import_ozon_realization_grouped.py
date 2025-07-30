# import_ozon_realization_grouped.py

import warnings
from collections import defaultdict

warnings.filterwarnings("ignore", category=UserWarning)
import requests
import pandas as pd
import glob
import os            # ← импорт был здесь
import re
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


def normalize_offer_id(val) -> str:
    """Remove trailing size like '-54' from seller article."""
    if val is None:
        return ""
    s = str(val).strip()
    return re.sub(r"-\d+$", "", s)



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
BASE_DIR   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(BASE_DIR, 'excel', 'Finmodel.xlsm')

  # подстрой под себя!
ORG_SHEET = 'НастройкиОрганизаций'
CFG_SHEET = 'Настройки'

def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def load_settings(wb):
    """Читаем настройки прямо из открытой книги (Book.caller())."""
    # ---------- 1. Организации -------------------------------------
    sht_orgs = wb.sheets[ORG_SHEET]
    df_all = sht_orgs.range('A1').expand().options(
        pd.DataFrame, header=1, index=False).value

    org_rows = []
    for _, row in df_all.iterrows():
        if pd.isna(row['Организация']) or str(row['Организация']).strip() == '':
            break
        org_rows.append(row)
    if not org_rows:
        raise Exception("В таблице 'НастройкиОрганизаций' нет данных!")
    df_orgs = pd.DataFrame(org_rows)

    def safe_str(x):
        if pd.isna(x):
            return ''
        s = str(x).strip()
        return s[:-2] if s.endswith('.0') else s

    orgs = []
    for _, row in df_orgs.iterrows():
        org       = safe_str(row['Организация'])
        client_id = safe_str(row['Client-Id'])
        token     = safe_str(row['Token_Ozon'])
        print(f"[DEBUG] Организация: {org} | Client-Id: {client_id} | Api-Key: {token}")
        if org and client_id and token:
            orgs.append({'org': org, 'client_id': client_id, 'token': token})

    # ---------- 2. Период загрузки ---------------------------------
    sht_cfg = wb.sheets[CFG_SHEET]
    df_params = sht_cfg.range('A1').expand().options(
        pd.DataFrame, header=False, index=False).value

    period_start, period_end = None, None
    for _, row in df_params.iterrows():
        if str(row[0]).strip() == 'ПериодНачало':
            period_start = pd.to_datetime(row[1])
        if str(row[0]).strip() == 'ПериодКонец':
            period_end = pd.to_datetime(row[1])

    if period_start is None:
        raise Exception("Не задана дата ПериодНачало в листе 'Настройки'")
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
    wb, app, created = get_workbook()
    log(f"Старт скрипта. Excel: {wb.fullname}")

    # ---------- читаем параметры ----------------------------------
    orgs, period_start, period_end = load_settings(wb)
    periods = get_periods(period_start, period_end)

    log(f"Организаций для запроса: {len(orgs)}")
    log(f"Периоды: {period_start:%Y-%m} — {period_end:%Y-%m}")

    # ---------- блок инициализации словаря groups ----------
   
    groups = defaultdict(lambda: {
        'org': '', 'year': 0, 'month': 0,
        'offer_id': '', 'sku': '', 'barcode': '', 'name': '',
        # комиссии после всех скидок
        'commission_del': 0.0,
        'commission_ret': 0.0,
        # для корректной выручки и среднего процента
        'seller_price_sum': 0.0,
        'commission_ratio_sum': 0.0,   # сумма (ratio × qty)
        'qty_sum': 0,                  # общее кол-во
        # блоки delivery / return
        'del': defaultdict(float),
        'ret': defaultdict(float)
    })


    for org_info in orgs:
        org       = org_info['org']
        client_id = org_info['client_id']
        token     = org_info['token']

        for p in periods:
            log(f"Запрос: {org} {p['year']}-{p['month']}")
            rows = fetch_ozon_data(org, client_id, token, p['year'], p['month'])

            for r in rows:
                it    = r.get('item', {})
                deliv = r.get('delivery_commission') or {}
                ret   = r.get('return_commission')   or {}

                offer = str(it.get('offer_id', '')).strip()
                key = (
                    org, p['year'], p['month'],
                    offer, it.get('sku', ''), it.get('barcode', ''), it.get('name', '')
                )
                g = groups[key]

                # статические поля
                g['org'], g['year'], g['month'], g['offer_id'], g['sku'], g['barcode'], g['name'] = key

                # количество доставленных штук
                qty_del = deliv.get('quantity', 0) or 0

                # оборот по цене продавца
                g['seller_price_sum']     += (r.get('seller_price_per_instance', 0) or 0) * qty_del
                # для среднего процента комиссии
                g['commission_ratio_sum'] += (r.get('commission_ratio', 0) or 0) * qty_del
                g['qty_sum']              += qty_del

                # итоговые комиссии уже после скидок (rule 1 %)
                g['commission_del'] += deliv.get('commission', 0) or 0
                g['commission_ret'] += ret.get('commission', 0) or 0

                # остальные показатели
                for src, dct in zip(['del', 'ret'], [deliv, ret]):
                    for f_api, f_out in [
                        ('amount', 'amount'), ('bonus', 'bonus'), ('commission', 'commission'),
                        ('compensation', 'compensation'), ('price_per_instance', 'price'),
                        ('quantity', 'qty'), ('standard_fee', 'std_fee'),
                        ('bank_coinvestment', 'bank'), ('stars', 'stars'),
                        ('pick_up_point_coinvestment', 'pvz'), ('total', 'total')
                    ]:
                        g[src][f_out] += dct.get(f_api, 0) or 0

   
    # --- Подготовка итоговой таблицы ---------------------------------
    result = []
    for g in groups.values():
        sold_qty = g['del']['qty'] - g['ret']['qty']
        partner_payouts = (
            g['del']['bank'] + g['del']['pvz'] + g['del']['stars']
        - g['ret']['bank'] - g['ret']['pvz'] - g['ret']['stars']
        )
        revenue = g['del']['amount'] - g['ret']['amount'] + partner_payouts

        base_fee        = g['del']['std_fee'] - g['ret']['std_fee']
        bonused_points  = g['del']['bonus']   - g['ret']['bonus']
        raw_commission  = base_fee - bonused_points
        min_commission  = 0.01 * g['seller_price_sum']
        net_reward      = raw_commission if raw_commission >= min_commission else min_commission

        commission_ratio_avg = (
            g['commission_ratio_sum'] / g['qty_sum'] if g['qty_sum'] else 0
        )

        result.append([
            g['org'], g['year'], g['month'], g['offer_id'], g['sku'], g['barcode'], g['name'],
            g['seller_price_sum'], commission_ratio_avg,
            g['del']['amount'], g['del']['bonus'], g['del']['commission'], g['del']['compensation'],
            g['del']['price'], g['del']['qty'], g['del']['std_fee'],
            g['del']['bank'], g['del']['stars'], g['del']['pvz'], g['del']['total'],
            g['ret']['amount'], g['ret']['bonus'], g['ret']['commission'], g['ret']['compensation'],
            g['ret']['price'], g['ret']['qty'], g['ret']['std_fee'],
            g['ret']['bank'], g['ret']['stars'], g['ret']['pvz'], g['ret']['total'],
            sold_qty, revenue, partner_payouts,
            bonused_points, base_fee, net_reward
        ])

    df = pd.DataFrame(result, columns=OUTPUT_HEADERS)
    write_to_excel(df)



if __name__ == '__main__':
    main()
