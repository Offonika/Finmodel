# update_plan_sales_ozon.py

import os
import xlwings as xw
import pandas as pd
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')

SHEET_SETTINGS = 'Настройки'
SHEET_SEASON   = 'Сезонность'
SHEET_SALES    = 'ФинотчетыОзон'
SHEET_PRICES   = 'ЦеныОзон'
SHEET_PLAN     = 'ПланПродажОзон'
TABLE_NAME     = 'PlanOzonTable'
TABLE_STYLE    = 'TableStyleMedium7'

MONTHS_CNT = 12
MONTH_NAMES = [f'Мес.{str(i+1).zfill(2)}' for i in range(MONTHS_CNT)]
CURRENT_MONTH = datetime.now().month
CURRENT_YEAR = datetime.now().year 

import logging

# Настройка логирования
LOG_PATH = os.path.join(BASE_DIR, 'plan_sales_ozon.log')
logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    filemode='w'   # каждый раз перезаписывать лог (или 'a' для дописывания)
)


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

def read_df(ws, idx_needed=()):
    """Безопасное чтение листа в DataFrame + проверка колонок"""
    df = ws.range(1, 1).expand().options(
        pd.DataFrame,
        header=True,   # было header=1   !!!
        index=False
    ).value
    idx = {c.strip(): i for i, c in enumerate(df.columns)}
    for col in idx_needed:
        if col not in idx:
            raise ValueError(f'Колонка «{col}» не найдена')
    return df, idx


def main():
    print("=== Старт update_plan_sales_ozon ===")
    wb, app = get_workbook()
    print("→ Открыт файл:", wb.fullname)

    # Чтение периода расчёта базовых продаж из листа "НастройкиОрганизаций"
    settings_ws = wb.sheets[SHEET_SETTINGS]
    settings_df = settings_ws.range(1, 1).expand().options(
        pd.DataFrame, header=1, index_col=None
    ).value

    def find_setting(df, param):
        """
        Ищет строку с нужным параметром в индексе DataFrame
        и возвращает значение из колонки 'Значение'.
        Ищем сначала точное совпадение, затем по вхождению.
        """
        # точное совпадение
        if param in df.index:
            return df.loc[param, df.columns[0]]

        # частичное (на случай лишних пробелов)
        for idx in df.index.astype(str):
            if param.lower() in idx.lower():
                return df.loc[idx, df.columns[0]]
        return None



    period_from = find_setting(settings_df, 'Период с')
    period_to   = find_setting(settings_df, 'Период по')
    print('DEBUG:', period_from, period_to)  # для отладки, потом уберите
    print("=== DEBUG: settings_df ===")
    print(settings_df.to_string())

    if not period_from or not period_to:
        print('❌ Не найден период в настройках!')
        if app: app.quit()
        return

    period_from = pd.to_datetime(period_from, dayfirst=True, errors='coerce')
    period_to   = pd.to_datetime(period_to, dayfirst=True, errors='coerce')
    print(f'→ Период для base: {period_from:%d.%m.%Y} — {period_to:%d.%m.%Y}')
    # 1. Сезонность
    season_df, _ = read_df(wb.sheets[SHEET_SEASON])
    season_factors = {
        str(r.iloc[0]).strip(): [
            safe_float(r.iloc[i]) if i < len(r) else 1.0
            for i in range(1, MONTHS_CNT + 1)
        ]
        for _, r in season_df.iterrows()
    }
    print(f'→ Лист {SHEET_SEASON} считан: {len(season_df)} строк')

    # 2. Финотчёты
    
    try:
        sales_df = wb.sheets[SHEET_SALES] \
                    .range(1, 1).expand() \
                    .options(pd.DataFrame, header=1, index=False).value
        print(f'→ Лист {SHEET_SALES}: {len(sales_df)} строк')
    except Exception:
        print(f'❌ Нет листа {SHEET_SALES}')
        if app: app.quit()
        return

    need_cols = {'Организация', 'Артикул_поставщика', 'SKU',
                'Год', 'Месяц', 'Продано шт.'}
    missing = need_cols - set(sales_df.columns)
    if missing:
        print(f'❌ В {SHEET_SALES} нет колонок: {", ".join(missing)}')
        if app: app.quit()
        return

    # ► безопасные числа
    sales_df['Год']        = sales_df['Год'].apply(safe_float).astype(int)
    sales_df['Месяц']      = sales_df['Месяц'].apply(safe_float).astype(int)
    sales_df['Продано шт.'] = sales_df['Продано шт.'].apply(safe_float)
    sales_df['ДатаМесяца'] = pd.to_datetime(
            sales_df['Год'].astype(str) + '-' + sales_df['Месяц'].astype(str).str.zfill(2) + '-01'
        )
    # ► только текущий год
    sales_df = sales_df[sales_df['Год'] == CURRENT_YEAR]
    print('Уникальные значения "Месяц":', sorted(sales_df['Месяц'].unique()))

    debug_df = sales_df[
        (sales_df['Организация'] == 'ИП Закиров Р.Х.') &
        (sales_df['SKU'].apply(normalize_sku) == normalize_sku('1499960988'))
    ]
    print(debug_df[['Год', 'Месяц', 'Артикул_поставщика', 'SKU', 'Продано шт.']])

    pivot = (
        sales_df
        .pivot_table(index=['Организация', 'SKU'],
                    columns='Месяц',
                    values='Продано шт.',
                    aggfunc='sum',
                    fill_value=0)
        .reindex(columns=range(1, 13), fill_value=0)
    )

    qty_map = {}
    sku_to_offer = {}

    for (org, sku), row in pivot.iterrows():
        key = (str(org).strip(), normalize_sku(sku))
        qty_map[key] = row.tolist()
        df_found = sales_df[
            (sales_df['Организация'] == org) &
            (sales_df['SKU'].apply(normalize_sku) == normalize_sku(sku))
        ]
        offer = str(df_found['Артикул_поставщика'].iloc[0]) if not df_found.empty else ''
        sku_to_offer[key] = offer.strip()





    # 3. Цены
    prices_df, _ = read_df(
        wb.sheets[SHEET_PRICES],
        idx_needed=('Артикул', 'Цена продавца с акциями')
    )
    price_map = {
        str(r['Артикул']).strip(): safe_float(r['Цена продавца с акциями'])
        for _, r in prices_df.iterrows()
    }
    print(f'→ Лист {SHEET_PRICES} считан: {len(prices_df)} строк')

    # 4. План
    rows = []
    logging.info("==== Начат расчет плана продаж Ozon ====")

    for (org, sku), hist in qty_map.items():
                # Отбор строк по организации, SKU и периоду
        org_mask = sales_df['Организация'] == org
        sku_mask = sales_df['SKU'].apply(normalize_sku) == normalize_sku(sku)
        period_mask = (sales_df['ДатаМесяца'] >= period_from) & (sales_df['ДатаМесяца'] <= period_to)
        mask = org_mask & sku_mask & period_mask
        df_period = sales_df[mask].sort_values('ДатаМесяца')

        done_months = df_period['Продано шт.'][df_period['Продано шт.'] > 0].tolist()

        msg_head = f"{org} | {sku_to_offer[(org, sku)]} | {sku} | Месяцы: {hist}"
        if not done_months:
            logging.warning(f"{msg_head} | Нет фактических продаж за {CURRENT_MONTH} месяцев — пропускаем.")
            continue
        base    = round(sum(done_months) / len(done_months)) if done_months else 0

        factors = season_factors.get(sku, [1.0] * MONTHS_CNT)
        price   = price_map.get(sku_to_offer[(org, sku)], 0.0)
        plan = []
        for i in range(MONTHS_CNT):
            if i < CURRENT_MONTH - 1:
                plan.append(round(hist[i]))
            else:
                plan.append(round(base * factors[i]))

        row_info = (f"{msg_head} | Базовое: {base} | План: {plan}")
        # Добавляем отметку, если май = 0
        if plan[4] == 0:
            logging.error(f"{row_info} | ВНИМАНИЕ: Месяц 5 (Май) = 0!")
        else:
            logging.info(row_info)
        if sum(plan) == 0:
            continue
        rows.append([org, sku_to_offer[(org, sku)], sku, base, price, *plan])


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
            plan_ws.api.Move(Before=wb.sheets[13].api)
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
    """Возвращает (wb, app).
    Если книга уже открыта в Excel — берём её,
    иначе создаём невидимый экземпляр и открываем файл."""
    try:
        wb = xw.Book.caller()
        print('→ Запуск из Excel-макроса')
        return wb, None
    except Exception:
        pass

    # Проверяем среди открытых книг только те, что имеют существующий путь
    for app in xw.apps:
        for bk in app.books:
            try:
                if bk.fullname \
                   and os.path.exists(bk.fullname) \
                   and os.path.samefile(bk.fullname, EXCEL_PATH):
                    print('→ Используем уже открытую книгу')
                    return bk, None
            except Exception:
                continue  # пропускаем книги без файла

    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(EXCEL_PATH, update_links=False)
    print('→ Книга была закрыта, открыли новую копию')
    return wb, app

logging.info("==== Расчет завершён ====")


if __name__ == '__main__':
    main()
