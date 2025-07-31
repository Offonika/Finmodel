from pathlib import Path
import xlwings as xw
import requests
import time
from datetime import datetime

# --- Константы ---
WB_PRICE_URL = 'https://discounts-prices-api.wildberries.ru/api/v2/list/goods/filter'
PAGE_LIMIT = 1000
PRICES_PAGE_PAUSE_SEC = 0.4
PRICES_MAX_RETRIES_FETCH = 3

HEADER_DICT = {
    'org': 'Организация',
    'nmID': 'Артикул_WB',
    'vendorCode': 'Артикул_поставщика',
    'sizeID': 'ID размера',
    'techSizeName': 'Размер',
    'price': 'Цена, ₽',
    'discountedPrice': 'Цена со скидкой, ₽',
    'clubDiscountedPrice': 'Клубная цена, ₽'
}
HEADERS_RU = list(HEADER_DICT.values())

EXCEL_PATH = Path(__file__).resolve().parents[1] / 'Finmodel.xlsm'

def log(msg):
    now = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    print(f"{now} {msg}")

def get_org_tokens(settings_sheet):
    data = settings_sheet.range('A1').expand('table').value
    if not data or len(data) < 2:
        log("❌ Нет данных в листе настроек организаций!")
        return []
    header = data[0]
    idx = {str(h).strip(): i for i, h in enumerate(header)}
    org_tokens = []
    for row in data[1:]:
        org = str(row[idx['Организация']]).strip() if row[idx['Организация']] else ''
        token = str(row[idx['Token_WB']]).strip() if row[idx['Token_WB']] else ''
        if org and token:
            log(f"[DEBUG] Организация: {org} | TokenWB: {token[:6]}...")
            org_tokens.append((org, token))
    log(f"[LOG] Всего организаций: {len(org_tokens)}")
    return org_tokens

def safe_fetch(url, headers):
    for attempt in range(PRICES_MAX_RETRIES_FETCH):
        try:
            resp = requests.get(url, headers=headers, timeout=20)
            if resp.status_code == 200:
                return resp.json()
            else:
                log(f"⚠️ Неудачный статус {resp.status_code} при запросе: {url}")
        except Exception as e:
            log(f"⚠️ Ошибка запроса: {e}")
        time.sleep(1.5)
    log(f"❌ Не удалось получить данные после {PRICES_MAX_RETRIES_FETCH} попыток: {url}")
    return {'data': {'listGoods': []}}

def get_or_create_sheet_with_header(wb, name, header):
    try:
        sh = wb.sheets[name]
        sh.clear()
        sh.range((1, 1)).value = header
        sh.range((1, 1), (1, len(header))).api.Font.Bold = True
    except Exception:
        sh = wb.sheets.add(name)
        sh.range((1, 1)).value = header
        sh.range((1, 1), (1, len(header))).api.Font.Bold = True
    return sh

def autofit_columns(sheet, cols_count):
    sheet.range((1, 1), (1, cols_count)).columns.autofit()

def load_wb_prices_by_size_xlwings(wb=None):
    created = False
    if wb is None:
        try:
            wb = xw.Book.caller()
            log('Запуск из Excel.')
        except Exception:
            wb = xw.Book(EXCEL_PATH)
            created = True
            log(f'Открыли файл: {EXCEL_PATH}')
    try:
        settings = wb.sheets['НастройкиОрганизаций']
    except Exception:
        log("❌ Нет листа 'НастройкиОрганизаций'!")
        if created:
            wb.close()
        return
    org_tokens = get_org_tokens(settings)
    if not org_tokens:
        log("❌ Нет организаций с токенами в настройках!")
        if created:
            wb.close()
        return
    output_sh = get_or_create_sheet_with_header(wb, 'Цены_WB', HEADERS_RU)
    try:
        output_sh.api.Tab.Color = 142661105
        log("→ Цвет ярлыка #84F8EA установлен")
    except Exception as e:
        log(f"⚠️ Не удалось установить цвет ярлыка: {e}")

    row_idx = 2
    any_data = False
    for org, token in org_tokens:
        offset = 0
        headers = {"Authorization": token}
        log(f'→ Организация: {org}')
        total_rows = 0
        while True:
            url = f"{WB_PRICE_URL}?limit={PAGE_LIMIT}&offset={offset}"
            resp = safe_fetch(url, headers)
            goods = resp.get('data', {}).get('listGoods', [])
            if not goods:
                log(f"  Конец данных offset={offset}")
                break
            rows = []
            for g in goods:
                for s in g.get('sizes', []):
                    row = [
                        org,
                        str(g.get('nmID')),
                        g.get('vendorCode'),
                        s.get('sizeID'),
                        s.get('techSizeName'),
                        s.get('price'),
                        s.get('discountedPrice'),
                        s.get('clubDiscountedPrice')
                    ]
                    rows.append(row)
            if rows:
                output_sh.range((row_idx, 1)).value = rows
                row_idx += len(rows)
                total_rows += len(rows)
                any_data = True
                log(f"  Получено строк: {len(rows)}, offset={offset}")
            else:
                log(f"  Нет новых строк offset={offset}")
            offset += PAGE_LIMIT
            time.sleep(PRICES_PAGE_PAUSE_SEC)
        log(f"  Всего выгружено строк для {org}: {total_rows}")

    autofit_columns(output_sh, len(HEADERS_RU))

    # --- Удалить предыдущую умную таблицу, если есть ---
    try:
        for tbl in output_sh.tables:
            if tbl.name == "WbPricesBySizeTable":
                tbl.api.Delete()
    except Exception as e:
        log(f"⚠️ Не удалось удалить старую таблицу: {e}")

    # --- Создать умную таблицу TableStyleLight1 ---
# ---
#     last_row = output_sh.range('A1').end('down').row
#     last_col = len(HEADERS_RU)
#     tbl_range = output_sh.range((1, 1), (last_row, last_col))
#     try:
#         output_sh.tables.add(tbl_range, name="WbPricesBySizeTable", table_style_name="TableStyleLight1", has_headers=True)
#         log("→ Умная таблица создана (TableStyleLight1)")
#     except Exception as e:
#         log(f"⚠️ Не удалось создать умную таблицу: {e}")

#     # --- Шапка: белый фон и чёрный текст ---
#     try:
#         header_range = output_sh.range((1, 1), (1, last_col))
#         header_range.api.Interior.Color = 0xFFFFFF  # белый
#         header_range.api.Font.Color = 0x000000      # чёрный
#     except Exception as e:
#         log(f"⚠️ Не удалось окрасить шапку: {e}")

#     # --- Цвет ярлыка #84F8EA (BGR: 0xEAF884) ---
#     try:
#         output_sh.api.Tab.Color = 0xEAF884
#         log("→ Цвет ярлыка #84F8EA установлен")
#     except Exception as e:
#         log(f"⚠️ Не удалось установить цвет ярлыка: {e}")
# ---
    # --- Переместить лист на позицию 14 ---
    try:
        if output_sh.index != 28:
            output_sh.api.Move(Before=wb.sheets[28].api)
            log("→ Лист перемещён на позицию 14")
    except Exception as e:
        log(f"⚠️ Не удалось переместить лист: {e}")

    if any_data:
        log('Готово! Данные записаны.')
    else:
        log('⚠️ Данные не были получены, таблица пуста!')

    if created:
        wb.save()
        wb.close()
        log(f'Файл сохранён и закрыт: {EXCEL_PATH}')
# После записи всех строк:
    col_nmID = HEADER_DICT['nmID']
    col_idx = HEADERS_RU.index(col_nmID) + 1
    last_row = output_sh.range('A1').end('down').row
    output_sh.range((2, col_idx), (last_row, col_idx)).api.NumberFormat = "@"

def main():
    log("=== Старт загрузки цен WB по размерам ===")
    load_wb_prices_by_size_xlwings()
    log("=== Конец работы ===")

if __name__ == "__main__":
    main()
