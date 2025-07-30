# update_plan_sales.py

import xlwings as xw
import pandas as pd
from datetime import datetime
#  добавьте сразу после import-ов
import argparse
from pathlib import Path
from scripts.sheet_utils import apply_sheet_settings

def parse_cli():
    p = argparse.ArgumentParser(add_help=False)
    p.add_argument('-f', '--file', help='Путь к Finmodel.xlsm')
    return p.parse_known_args()[0].file

CLI_XLSM = parse_cli()                # None, если параметр не передали

EXCEL_PATH = str(Path(CLI_XLSM)) if CLI_XLSM else str(Path(__file__).resolve().parents[1] / 'Finmodel.xlsm')

SHEET_SETTINGS   = 'Настройки'
SHEET_PRODUCTS   = 'Номенклатура_WB'
SHEET_FACTS      = 'ФинотчетыWB'
SHEET_SEASON     = 'Сезонность'
SHEET_PRICES     = 'Цены_WB'
SHEET_PLAN       = 'План_ПродажWB'

TABLE_NAME = 'PlanSalesTable'
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

def parse_date(val):
    if isinstance(val, datetime):
        return val
    if isinstance(val, (float, int)) and not pd.isna(val):
        try:
            return datetime(1899, 12, 30) + pd.to_timedelta(int(val), unit='D')
        except Exception:
            pass
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(val), fmt)
        except Exception:
            continue
    return pd.NaT

def safe_float(val):
    if pd.isna(val):
        return 1
    if isinstance(val, str):
        val = val.replace(',', '.')
    try:
        return float(val)
    except Exception:
        return 1

def normalize_artwb(val):
    if pd.isna(val):
        return ""
    s = str(val)
    if s.endswith('.0'):
        s = s[:-2]
    return s.strip()

def norm_key(artwb):
    return normalize_artwb(artwb)

def clean_org(s):
    """
    Универсальная очистка строки от пробелов, неразрывных пробелов и скрытых символов.
    """
    return str(s).replace('\xa0', ' ').replace('\u200b', '').replace('\t', '').strip()

def main():
    print("=== Старт update_plan_sales ===")
    debug_artwb = '173304613'  # nmID для отладки!
    key_for_debug = norm_key(debug_artwb)
    wb, app = get_workbook()
    try:
        settings_ws   = wb.sheets[SHEET_SETTINGS]
        prod_ws       = wb.sheets[SHEET_PRODUCTS]
        facts_ws      = wb.sheets[SHEET_FACTS]
        season_ws     = wb.sheets[SHEET_SEASON]
        price_ws      = wb.sheets[SHEET_PRICES]
        print('→ Листы найдены')
    except Exception as e:
        print('❌ Ошибка загрузки листов:', e)
        if app:
            app.quit()
        return

    # 1. Период усреднения
    settings = settings_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
    dt_from = dt_to = None
    for _, row in settings.iterrows():
        if str(row.iloc[0]).strip() == 'Период с':
            dt_from = parse_date(row.iloc[1])
        if str(row.iloc[0]).strip() == 'Период по':
            dt_to = parse_date(row.iloc[1])
    if not dt_from or not dt_to:
        print('❌ Не заданы "Период с/по" в Настройках')
        if app:
            app.quit()
        return
    months_cnt = (dt_to.year-dt_from.year)*12 + (dt_to.month-dt_from.month) + 1
    print(f'→ Период: {dt_from.strftime("%d.%m.%Y")} - {dt_to.strftime("%d.%m.%Y")} ({months_cnt} мес.)')

    # 2. Сезонные коэффициенты
    season = season_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
    season_factors = {}
    for _, row in season.iterrows():
        key = row.iloc[0]
        vals = [safe_float(row.iloc[i]) for i in range(1, 13)]
        season_factors[key] = vals
    print(f'→ Считано строк сезонности: {len(season_factors)}')

    # 3. Чтение данных и унификация идентификаторов
   # --- Чтение данных и очистка колонок ---
    products = prod_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value

# ВСТАВЬ СЮДА ↓↓↓↓↓
    if 'Артикул_WB' in products.columns:
        products['Артикул_WB'] = products['Артикул_WB'].apply(
            lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x).strip()
        )

    products.columns = products.columns.str.strip()
# Сразу после products = ...
    print('[DEBUG] ВСЯ выгрузка Артикул_WB (первые 50):')
    print(products['Артикул_WB'].head(50).to_list())

    # Посчитать сколько строк с нужным артикулом
    print('[DEBUG] Кол-во строк с 173304613:', (products['Артикул_WB'] == '173304613').sum())

    # КОРРЕКЦИЯ Артикул_WB: float -> str
    if 'Артикул_WB' in products.columns:
        products['Артикул_WB'] = products['Артикул_WB'].apply(lambda x: str(int(x)) if pd.notna(x) and float(x).is_integer() else str(x))

    print('[DEBUG] Колонки products:', list(products.columns))
    if not products.empty:
        print('[DEBUG] Пример первой строки products:', products.iloc[0].to_dict())


    # Отладочный вывод: список колонок и пример первой строки
    print('[DEBUG] Колонки products:', list(products.columns))
    if not products.empty:
        print('[DEBUG] Пример первой строки products:', products.iloc[0].to_dict())

    prices   = price_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
    facts    = facts_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
# === ВЫВОД ВСЕХ НЕЧИСЛОВЫХ ЗНАЧЕНИЙ В 'Итого_продано' ===
    # === ВЫВОД ВСЕХ НЕЧИСЛОВЫХ ЗНАЧЕНИЙ В 'Итого_продано' ===
    print("\n=== Поиск нечисловых значений в 'Итого_продано' ===")
    bad_count = 0
    for i, val in enumerate(facts['Итого_продано']):
        try:
            # Попытка привести к числу (float)
            _ = float(str(val).replace(',', '.').replace(' ', '').replace('₽', ''))
        except Exception:
            print(f"[BAD] строка #{i}: {val!r} (type={type(val)})")
            print("  Вся строка:", facts.iloc[i].to_dict())
            bad_count += 1
    print(f"=== Найдено нечисловых значений: {bad_count} ===\n")

    # --- Удаляем лишние пробелы в заголовках! ---
    products.columns = products.columns.str.strip()
    prices.columns   = prices.columns.str.strip()
    facts.columns    = facts.columns.str.strip()
    print(f"[DEBUG] Колонки products: {list(products.columns)}")
    print(f"[DEBUG] Колонки facts: {list(facts.columns)}")
    print(f"[DEBUG] Колонки prices: {list(prices.columns)}")


    # --- Унификация идентификаторов ---
    if 'Код_номенклатуры' in facts.columns:
        facts['Артикул_WB'] = facts['Код_номенклатуры'].apply(normalize_artwb)
    if 'Артикул_WB' not in products.columns and 'nmID' in products.columns:
        products['Артикул_WB'] = products['nmID'].apply(normalize_artwb)
    if 'Артикул_WB' not in prices.columns and 'nmID' in prices.columns:
        prices['Артикул_WB'] = prices['nmID'].apply(normalize_artwb)

    # Индексы
    idx_prod = {h: i for i, h in enumerate(products.columns)}

    # 4. Подготовка списка товаров (ключ: Артикул_WB)
    prod_list = []
    for _, r in products.iterrows():
        prod_list.append({
            'org':    clean_org(r.iloc[idx_prod['Организация']]),
            'artwb':  normalize_artwb(r.iloc[idx_prod['Артикул_WB']]),
            'vendor': r.iloc[idx_prod['Артикул_поставщика']],
            'subject': r.iloc[idx_prod['Предмет']],
        })


    # 5. Примеры ключей для отладки
    prod_keys = [str(norm_key(p['artwb'])).strip() for p in prod_list[:1000]]
    fact_keys = [str(norm_key(row['Артикул_WB'])).strip() for _, row in facts.iterrows()]
    print(f"=== Пример prod_keys: {prod_keys[:30]}")
    print(f"=== Пример fact_keys: {fact_keys[:30]}")
    print(f"Ключ '173304613' в prod_keys: {'173304613' in prod_keys}")
    print(f"Ключ '173304613' в fact_keys: {'173304613' in fact_keys}")


    # Проверяем наличие debug-ключа в prod_list и facts
    debug_key = key_for_debug
    found_in_prod = debug_key in prod_keys
    found_in_facts = debug_key in fact_keys
    print(f"\n=== Поиск ключа {debug_key}: ===")
    print(f"- В prod_list: {'НАЙДЕН' if found_in_prod else 'НЕТ'}")
    print(f"- В facts: {'НАЙДЕН' if found_in_facts else 'НЕТ'}")

    # Ищем ближайшие совпадения по nmID для диагностики (по nmID без организации)
    debug_nmID = debug_key
    prod_similar = [k for k in prod_keys if k == debug_nmID]
    fact_similar = [k for k in fact_keys if k == debug_nmID]
    print(f"\n=== Ключи из prod_list с этим nmID: {prod_similar}")
    print(f"=== Ключи из facts с этим nmID: {fact_similar}")

    # Если не найдено — выводим все уникальные nmID в prod_list и facts для ручной сверки
    prod_nmid_set = set(prod_keys)
    fact_nmid_set = set(fact_keys)
    print(f"\n=== nmID в prod_list (первые 20): {list(prod_nmid_set)[:20]}")
    print(f"=== nmID в facts (первые 20): {list(fact_nmid_set)[:20]}")

    # 6. Год анализа — всегда текущий год!
    now = datetime.now()
    year_plan = now.year
    current_month = now.month
    print(f"→ Будет использоваться год анализа: {year_plan}")

    # 7. Сбор продаж по месяцам (факты только текущий год)
    col_artwb  = 'Артикул_WB'
    col_date   = 'Дата'
    col_sold   = 'Итого_продано'
    qty_map = {norm_key(p['artwb']): [0]*12 for p in prod_list}
    used_facts = 0

    def safe_num(val):
        try:
            if pd.isna(val):
                return 0
            if isinstance(val, (int, float)):
                return val
            return float(str(val).replace(',', '.').replace(' ', '').replace('₽', ''))
        except Exception:
            return None  # Вернём None, чтобы отследить ошибку

    for idx, row in facts.iterrows():
        d = parse_date(row[col_date])
        key = norm_key(row[col_artwb])
        if pd.isna(d) or d.year != year_plan:
            continue
        if key not in qty_map:
            continue

        sold_raw = row[col_sold]
        sold = safe_num(sold_raw)
        if sold is None:
            print(f"[BAD SOLD] Строка #{idx}: key={key} дата={row[col_date]} sold_raw={sold_raw!r} type={type(sold_raw)}")
            print(f"Строка целиком: {row.to_dict()}")
            continue  # пропускаем ошибочную строку
        
        qty_map[key][d.month-1] += sold
        used_facts += 1

    # 8. Цены
    price_map = {}
    if 'Организация' in prices.columns:
        for _, row in prices.iterrows():
            org = clean_org(row['Организация'])
            key = norm_key(row['Артикул_WB'])
            price = row.get('Цена со скидкой, ₽', 0)
            price_map[(org, key)] = float(str(price).replace(',', '.')) if pd.notna(price) else 0

    else:
        for _, row in prices.iterrows():
            key = norm_key(row['Артикул_WB'])
            price = row.get('Цена со скидкой, ₽', 0)
            price_map[key] = float(str(price).replace(',', '.')) if pd.notna(price) else 0


    # 9. Группировка по фактам (только текущий год!)
    facts_this_year = facts.copy()
    facts_this_year['__parsed_date'] = facts_this_year[col_date].apply(parse_date)
    facts_this_year = facts_this_year[facts_this_year['__parsed_date'].apply(lambda d: d.year == year_plan if pd.notna(d) else False)]
    facts_this_year['__key'] = facts_this_year[col_artwb].apply(normalize_artwb)
    facts_this_year['__month'] = facts_this_year['__parsed_date'].apply(lambda d: d.month if pd.notna(d) else 0)

    def safe_num(val):
        try:
            if pd.isna(val):
                return 0.0
            if isinstance(val, (int, float)):
                return val
            return float(str(val).replace(',', '.').replace(' ', '').replace('₽', ''))
        except Exception:
            return 0.0

    facts_this_year[col_sold] = facts_this_year[col_sold].apply(safe_num)

    grouped_facts = facts_this_year.groupby(['__key', '__month'])[col_sold].sum().to_dict()

    print('\n=== Пример ключей grouped_facts (первые 20):')
    for k in list(grouped_facts.keys())[:20]:
        print(f'{k}: {grouped_facts[k]}')

    print(f'\n=== Продажи по ключу {key_for_debug}:')
    for m in range(1, 13):
        val = grouped_facts.get((key_for_debug, m), None)
        print(f'  Месяц {m:02d}: {val}')

    # 10. Формирование итоговой таблицы (логика: факт до текущего месяца, далее — прогноз)
    rows = []
    
    for p in prod_list:
        org = clean_org(p['org'])
        key = norm_key(p['artwb'])
        if (org, key) in price_map:
            price = round(price_map[(org, key)])
        elif key in price_map:
            price = round(price_map[key])
        else:
            price = 0
        


        arr = [int(grouped_facts.get((key, m), 0)) for m in range(1, 13)]
        base = round(sum(qty_map[key]) / months_cnt) if months_cnt else 0
        factors = season_factors.get(p['subject'], [1]*12)
        month_vals = []
      
        for i in range(12):
            month_num = i + 1
            if month_num < current_month:
                # Прошедшие месяцы: только факт, иначе 0
                month_val = arr[i] if arr[i] > 0 else 0
            else:
                # Текущий и будущие месяцы: только прогноз
                month_val = round(base * factors[i])
            month_vals.append(month_val)


        if key == key_for_debug:
            print(f'\n=== Итоговый arr для {key}: {arr}')
            print(f'=== Плановые значения для {key}:')
            for i in range(12):
                print(f'  Месяц {i+1:02d}: {month_vals[i]} (факт={arr[i]}, base={base}, фактор={factors[i]})')
        total_row = sum(month_vals)
        if total_row == 0:
            continue
        line = [p['org'], p['artwb'], p['vendor'], p['subject'], base, price] + month_vals + [total_row]
        rows.append(line)
    rows.sort(key=lambda x: -x[4])

    print(f'→ Итого товаров с планом: {len(rows)}')

    # 11. Вывод в Excel
    try:
        plan_ws = wb.sheets[SHEET_PLAN]
        plan_ws.clear()
    except Exception:
        plan_ws = wb.sheets.add(SHEET_PLAN)

    apply_sheet_settings(wb, SHEET_PLAN)

    header = ['Организация','Артикул_WB','Артикул_поставщика','Предмет','Базовое кол-во','Плановая цена, ₽'] + \
             [f'Мес.{str(i+1).zfill(2)}' for i in range(12)] + ['Всего']

    plan_ws.range(1,1).value = header
    if rows:
        plan_ws.range(2, 1).value = rows
        last_row = len(rows) + 1
        total_col = len(header)
        for tbl in plan_ws.tables:
            if tbl.name == TABLE_NAME:
                tbl.delete()
        table_range = plan_ws.range((1,1), (last_row, total_col))
        plan_ws.tables.add(table_range, name=TABLE_NAME, table_style_name=TABLE_STYLE, has_headers=True)
        plan_ws.range('A1').expand().columns.autofit()
        plan_ws.api.Rows(1).Font.Bold = True
        plan_ws.api.Application.ActiveWindow.SplitRow = 1
        plan_ws.api.Application.ActiveWindow.FreezePanes = True
    else:
        print('Нет строк для вывода — таблица не создаётся')

    if app:
        wb.save()
        app.quit()
    print('=== Скрипт успешно завершён ===')

if __name__ == '__main__':
    main()
