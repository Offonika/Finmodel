# calculate_cogs_batched.py

import os
import xlwings as xw
import pandas as pd
import math
import logging, datetime, pathlib
RUS_TO_LAT = str.maketrans("АВЕКМНОРСТХ",
                           "ABEKMHOPCTX")  # кир → лат
# --- Вставь сразу после import'ов ---
def rgb_to_excel_tab_color(r, g, b):
    """Преобразует RGB в BGR-целое для ws.api.Tab.Color (Excel COM)."""
    return b + g * 256 + r * 256 * 256

def hex_to_excel_tab_color(hex_str):
    """HEX #RRGGBB в Tab.Color (BGR-целое)."""
    hex_str = hex_str.strip().lstrip('#')
    if len(hex_str) != 6:
        raise ValueError("HEX должен быть 6 символов, например, #92D050")
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return rgb_to_excel_tab_color(r, g, b)

def norm(key: str) -> str:
    """
    Нормализует артикул:
    • обрезает пробелы
    • приводит к верхнему регистру
    • переводит русские 'A,B, ...' в латиницу
    """
    if key is None:
        return ""
    return str(key).replace(" ", " ").strip().upper().translate(RUS_TO_LAT)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))   # ← был позже – перенесли сюда!
LOG_DIR  = pathlib.Path(BASE_DIR, "log")
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / f"cogs_{datetime.datetime.now():%Y%m%d_%H%M%S}.log"
# ---------- ДОБАВЬТЕ СРАЗУ ПОСЛЕ import'ов --------------------


LOG_FILE = os.path.join(
    BASE_DIR,
    f"planned_indicators_{datetime.datetime.now():%Y%m%d_%H%M%S}.log"
)
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,                 # INFO→видно всё; WARNING→только ошибки
    format="%(asctime)s  %(message)s",
    encoding="utf-8",
)
log = logging.getLogger(__name__)
# --------------------------------------------------------------

logging.basicConfig(
    filename=str(LOG_FILE),
    filemode="w",
    # INFO скрывает debug-сообщения
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger(__name__)

def log(msg, level="info"):
    getattr(logger, level)(msg)      # пишем в файл
    # если хотите видеть всё ещё и в консоли, раскомментируйте:
    # print(msg)
# --------------------------------
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')


# ЗАМЕНЁННЫЕ НАЗВАНИЯ ЛИСТОВ
SHEET_PRODUCTS = 'Номенклатура_WB'
SHEET_PRICES   = 'ЗакупочныеЦены'
SHEET_DUTIES   = 'ТаможенныеПошлины'
SHEET_SETTINGS = 'Настройки'
SHEET_ORGS = 'НастройкиОрганизаций'

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
def get_logistics_mode(org, orgs_ws):
    df = orgs_ws.range(1, 1).expand().options(pd.DataFrame, header=1, index=False).value
    row = df[df.iloc[:, 0] == org]
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
    log("=== Старт batch расчёта себестоимости ===")
    wb, app = get_workbook()           # app=None → запущено из Excel; иначе invis-Excel

    try:                               # ───── ВСЁ внутри этого try ─────
        # 1. Получаем нужные листы
        try:
            prod_ws     = wb.sheets[SHEET_PRODUCTS]
            price_ws    = wb.sheets[SHEET_PRICES]
            duty_ws     = wb.sheets[SHEET_DUTIES]
            orgs_ws     = wb.sheets[SHEET_ORGS]
            settings_ws = wb.sheets[SHEET_SETTINGS]
        except Exception as e:
            print(f"❌ Не найден один из листов: {e}")
            log(f"Критическая ошибка: {e}", "error")
            return                       # выходим, finally всё закроет

        # 2. Читаем глобальные параметры
        global_params = read_settings(settings_ws)
        log(f"→ Параметры: {global_params}")

        # 3. Загружаем таблицы в DataFrame
        prod_df  = prod_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
        price_df = price_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value
        duty_df  = duty_ws.range(1,1).expand().options(pd.DataFrame, header=1, index=False).value

        # 4. Строим словари для быстрого поиска
        price_dict = {norm(r['Артикул_поставщика']): r for _, r in price_df.iterrows()}
        duty_dict  = {str(r['Предмет']).strip(): r for _, r in duty_df.iterrows()}

        # --- ДОБАВЬ ЭТОТ БЛОК ---

        # Логируем список всех уникальных предметов из товаров и из пошлин
        unique_subjects_products = set(str(row['Предмет']).strip() for _, row in prod_df.iterrows())
        unique_subjects_duties   = set(str(row['Предмет']).strip() for _, row in duty_df.iterrows())

        log(f"[INFO] Всего уникальных предметов в товарах: {len(unique_subjects_products)}")
        log(f"[INFO] Всего уникальных предметов в пошлинах: {len(unique_subjects_duties)}")

        missing_in_duties = sorted(s for s in unique_subjects_products if s not in unique_subjects_duties)
        if missing_in_duties:
            log(f"[WARN] В пошлинах отсутствуют {len(missing_in_duties)} предметов из товаров. Примеры: {missing_in_duties[:10]}")
        else:
            log("[INFO] Все предметы из товаров найдены в пошлинах.")

        extra_in_duties = sorted(s for s in unique_subjects_duties if s not in unique_subjects_products)
        if extra_in_duties:
            log(f"[INFO] В пошлинах есть {len(extra_in_duties)} лишних предметов, которых нет в товарах. Примеры: {extra_in_duties[:10]}")


        
  # 5. Готовим лист результата
        result_ws = wb.sheets[SHEET_RESULT] if SHEET_RESULT in [s.name for s in wb.sheets] \
                    else wb.sheets.add(SHEET_RESULT)

        # --- Установка цвета ярлыка #92D050 (зелёный) ---
        try:
            result_ws.api.Tab.Color = hex_to_excel_tab_color("#92D050")  # <--- Вот так!
        except Exception as e:
            print(f"⚠️ Не удалось задать цвет ярлыка: {e}")

        # --- Перемещение листа на 11-ю позицию ---
        try:
            if result_ws.index != 11:
                result_ws.api.Move(Before=wb.sheets[10].api)
                print("→ Лист 'РасчётСебестоимости' перемещён на позицию 11")
        except Exception as e:
            print(f"⚠️ Не удалось переместить лист: {e}")

            
     

        header = ['Организация','Артикул_поставщика','Предмет','Наименование',
                  'Закуп_Цена_руб','Логистика_руб','Пошлина_руб','НДС_руб',
                  'Себестоимость_руб','Себестоимость_без_НДС_руб','Входящий_НДС_руб']
        result_ws.clear();  result_ws.range(1,1).value = header
        first_free = 2

        # 6. Основной цикл по товарам чанками
        skipped = 0
        for chunk_start in range(0, len(prod_df), BATCH_SIZE):
            chunk_end = min(len(prod_df), chunk_start + BATCH_SIZE)
            batch_out = []

            for i in range(chunk_start, chunk_end):
                row = prod_df.iloc[i]
                org          = row['Организация']
                vendor_orig  = row['Артикул_поставщика']
                vendor_norm  = norm(vendor_orig)
                subject      = row['Предмет']
                name         = row['Название']
                weight       = safe_float(row['Вес_брутто'])

                price_row = price_dict.get(vendor_norm)
                if price_row is None:
                    # было: log(..., "warning")
                    log(f"Skip {vendor_orig} ({org}) – нет закупочной цены", "debug")
                    skipped += 1
                    continue

                # --- расчёты ---
                price_val = safe_float(price_row.get('Закуп_Цена'))
                rate      = global_params['usdRate'] if price_row.get('Валюта') == 'USD' \
                         else global_params['cnyRate'] if price_row.get('Валюта') == 'CNY' else 1
                purchase_rub = price_val * rate

                duty_row       = duty_dict.get(subject)
                logistics_mode = get_logistics_mode(org, orgs_ws)

                kg_rate        = global_params['cargoRatePerKg'] if logistics_mode=='Карго' \
                            else global_params['whiteRatePerKg']
                logistics_rub  = weight * kg_rate * global_params['usdRate']

                duty_rate = 0
                if logistics_mode == 'Белая' and duty_row is not None:
                    raw = duty_row.get('Ставка_пошлины') or duty_row.get('Пошлина')
                    log(f"Пошлина для {subject}: raw={raw}", "info")
                    if raw:
                        raw_str = str(raw).replace('%','').replace(',','.').strip()
                        try:
                            if float(raw_str) < 1:   # Уже доля, например 0.142 → не делим
                                duty_rate = float(raw_str)
                            else:                   # Вдруг где-то 14.2 → делим
                                duty_rate = float(raw_str) / 100
                        except Exception as e:
                            log(f"[ERROR] Не удалось привести raw='{raw}' к числу: {e}", "error")
                duty_rub = purchase_rub * duty_rate


                vat_rub  = (purchase_rub + duty_rub + logistics_rub) * global_params['ndsRateWhite'] \
                           if logistics_mode == 'Белая' else 0

                total_cogs      = purchase_rub + duty_rub + logistics_rub + vat_rub
                cogs_without_vat = total_cogs - vat_rub

                batch_out.append([
                    org, vendor_orig, subject, name,
                    round(purchase_rub), round(logistics_rub), round(duty_rub),
                    round(vat_rub),      round(total_cogs),    round(cogs_without_vat),
                    round(vat_rub)
                ])

            if batch_out:
                result_ws.range((first_free, 1)).value = batch_out
                first_free += len(batch_out)
                log(f"добавлено строк: {len(batch_out)}")

        log(f"Расчёт завершён. Итоговых строк: {first_free-2}, пропущено без цены: {skipped}")
        #print(f"✓ COGS рассчитан: {first_free-2} строк, пропусков {skipped}")

        # 7. Оформляем умную таблицу
        for tbl in result_ws.tables:
            if tbl.name == TABLE_NAME:
                tbl.delete()

        rng = result_ws.range((1,1), (first_free-1, len(header)))
        result_ws.tables.add(rng, name=TABLE_NAME,
                             table_style_name=TABLE_STYLE, has_headers=True)
        result_ws.autofit()
        log("Готово, файл сохранён"); print("✓ Готово!")

    finally:                           # ───── закрываем Excel, если нужен ─────
        if app is not None:
            wb.save(); wb.close()
            app.quit(); del app
            log("Excel закрыт корректно")

# ------------------------------------------
if __name__ == "__main__":
    main()
