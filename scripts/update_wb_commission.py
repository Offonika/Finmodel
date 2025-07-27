# update_wb_commission.py

import os
import xlwings as xw
import requests
from scripts.sheet_utils import apply_sheet_settings

# ==== КОНСТАНТЫ ====
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'excel', 'Finmodel.xlsm')
SETTINGS_SHEET = 'НастройкиОрганизаций'
TARGET_SHEET = 'КомиссияWB'
HEADERS = ['Parent Category', 'Subject Name', 'Commission, %']
API_URL = 'https://common-api.wildberries.ru/api/v1/tariffs/commission?locale=ru'

def get_workbook():
    try:
        wb = xw.Book.caller()  # Запуск из Excel (макрос RunPython)
        app = None
        print('→ Запуск из Excel через макрос')
    except Exception:
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_PATH)
        print(f'→ Запуск из консоли, открыт файл: {EXCEL_PATH}')
    return wb, app

def get_idx(header_row):
    return {h.strip(): i for i, h in enumerate(header_row)}

def main():
    print("=== Старт обновления комиссии WB ===")
    wb, app = get_workbook()
    try:
        # --- Получаем лист с настройками ---
        if SETTINGS_SHEET not in [s.name for s in wb.sheets]:
            raise Exception(f'❌ Лист "{SETTINGS_SHEET}" не найден!')

        sht_set = wb.sheets[SETTINGS_SHEET]
        cfgHdr = sht_set.range('A1').expand('right').value
        idx = get_idx(cfgHdr)

        if 'Token_WB' not in idx:
            raise Exception('❌ Нет колонки Token_WB в шапке!')

        token = sht_set.range(2, idx['Token_WB'] + 1).value
        if not token or not isinstance(token, str):
            raise Exception('❌ В первой строке нет Token_WB!')

        print('→ Токен найден, делаем запрос к WB API...')

        # --- Запрос к API ---
        resp = requests.get(API_URL, headers={'Authorization': token.strip()})
        if resp.status_code != 200:
            raise Exception(f"❌ Ошибка запроса WB API: {resp.status_code} – {resp.text}")

        data = resp.json()
        report = data.get('report')
        if not isinstance(report, list):
            raise Exception('❌ Неожиданный формат ответа, нет массива report[]')

        # --- Формируем данные для записи ---
        out = [HEADERS]
        for item in report:
            pct = float(item.get('kgvpMarketplace') or 0) / 100
            out.append([item.get('parentName', ''), item.get('subjectName', ''), pct])
        print(f"→ Загружено строк: {len(out) - 1}")

        # --- Создаём/очищаем целевой лист ---
        
        sheets = wb.sheets
        n_target = 30

        if TARGET_SHEET not in [s.name for s in sheets]:
            # Если меньше 30 листов — добавим в конец, иначе на 30‑ю позицию
            if len(sheets) < n_target:
                sht_tar = sheets.add(TARGET_SHEET, after=sheets[-1])
                print(f'→ Лист "{TARGET_SHEET}" создан в конце (меньше 30 листов)')
            else:
                sht_tar = sheets.add(TARGET_SHEET, before=sheets[n_target-1])
                print(f'→ Лист "{TARGET_SHEET}" создан на позиции {n_target}')
        else:
            sht_tar = sheets[TARGET_SHEET]
            sht_tar.clear()
            print(f'→ Лист "{TARGET_SHEET}" очищен')

        apply_sheet_settings(wb, TARGET_SHEET)

        sht_tar.range('A1').value = out

        # Форматирование процентов
        last_row = len(out)
        if last_row > 1:
            sht_tar.range((2, 3), (last_row, 3)).number_format = '0.00%'
        print("→ Данные записаны и отформатированы")

        wb.save()
        print("✅ Обновление завершено")
    except Exception as e:
        print(f'❌ Ошибка: {e}')
    finally:
        if app is not None:
            app.quit()

if __name__ == '__main__':
    main()
