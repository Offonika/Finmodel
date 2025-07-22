# update_monthly_scenario_calc.py

import os
import xlwings as xw
import win32com.client
import math#
import re
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')

# Все названия листов вынесены в словарь
SHEET_NAMES = {
    'plan_sales': 'План_ПродажWB',
    'plan_rev': 'План_ВыручкиWB',
    'dict': 'Номенклатура_WB',
    'comm': 'КомиссияWB',
    'cfg': 'Настройки',
    'cost': 'РасчётСебестоимости',
    'result': 'РасчётЭкономикиWB'
}

MONTHS = [f'Мес.{str(i+1).zfill(2)}' for i in range(12)]

def wb_code_key(val):
    """Артикул WB приводится к строке целого числа (убирает .0!)"""
    try:
        return str(int(float(val)))
    except Exception:
        return str(val).strip()

def get_workbook():
    try:
        wb = xw.Book.caller()
        app = None
        from_caller = True
    except Exception:
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_PATH)
        from_caller = False
    return wb, app, from_caller

def idx_from_header(header_row):
    """Строит словарь индексов колонок по заголовкам"""
    return {str(h).strip(): i for i, h in enumerate(header_row)}

def col_letter(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def to_num(val):
    try:
        if isinstance(val, (float, int)):
            return val
        val = str(val).replace('₽', '').replace('\xa0', '').replace(' ', '').replace(',', '.')
        return float(val) if val else 0
    except Exception:
        return 0

def get_workbook():
    try:
        wb = xw.Book.caller()
        app = None
        from_caller = True
        print(f'→ Запуск из Excel (RunPython): {EXCEL_PATH}')
    except Exception:
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_PATH)
        from_caller = False
        print(f'→ Запуск из консоли, открыт файл: {EXCEL_PATH}')
    return wb, app, from_caller



def norm_key(val):
    """Убирает все виды пробелов и приводит к верхнему регистру"""
    return re.sub(r'\s+', '', str(val)).upper()


def main():
    import time
    start = time.time()
    print('⏳ [START] Скрипт расчёта начал выполнение')

    wb, app, from_caller = get_workbook()
    try:
        # 1. Проверяем/создаём необходимые листы
        available_sheets = {s.name: s for s in wb.sheets}
        sh = {}
        for key, sheet_name in SHEET_NAMES.items():
            if sheet_name in available_sheets:
                sh[key] = available_sheets[sheet_name]
            elif key == 'cost':
                sh[key] = wb.sheets.add(sheet_name)
                print(f'➕ Лист "{sheet_name}" создан (РасчётСебестоимости)')
            elif key == 'result':
                sh[key] = wb.sheets.add(sheet_name)
                print(f'➕ Лист "{sheet_name}" создан (итоговый)')
            else:
                print(f'❌ Не найден лист: {sheet_name}')
                return
        # --- Загрузка % выкупа из %ВыкупаWB ---
        print('📈 Загрузка процента выкупа из %ВыкупаWB...')
        wb_sheetnames = [s.name for s in wb.sheets]
        if '%ВыкупаWB' in wb_sheetnames:
            s_wb = wb.sheets['%ВыкупаWB']
            wb_table = s_wb.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
            redemption_rate = {str(row['nmId']): float(row['% выкупа']) for _, row in wb_table.iterrows() if not pd.isna(row['nmId']) and not pd.isna(row['% выкупа'])}
        else:
            print('[WARN] Лист %ВыкупаWB не найден! Будет использоваться 95% по умолчанию.')
            redemption_rate = {}

        # 2. Индексы по заголовкам
        print('📄 Загрузка заголовков...')
        pIdx = idx_from_header(sh['plan_sales'].range(1, 1).expand('right').value)
        rIdx = idx_from_header(sh['plan_rev'].range(1, 1).expand('right').value)
        dIdx = idx_from_header(sh['dict'].range(1, 1).expand('right').value)
        cIdx = idx_from_header(sh['comm'].range(1, 1).expand('right').value)
        sIdx = idx_from_header(sh['cost'].range(1, 1).expand('right').value)

        # 3. Справочник товаров — используем Артикул_WB как ключ!
        print('📘 Чтение справочника товаров...')
        dicts = {}
        for r in sh['dict'].range(2, 1).expand('table').value:
            wb_code = wb_code_key(r[dIdx['Артикул_WB']])
            # пробуем взять объем из колонки
            try:
                volL = float(r[dIdx.get('Объем_литр', -1)]) if 'Объем_литр' in dIdx else None
            except Exception:
                volL = None
            if not volL or volL == 0:
                try:
                    width = float(r[dIdx['Ширина']]) if 'Ширина' in dIdx else 0
                    height = float(r[dIdx['Высота']]) if 'Высота' in dIdx else 0
                    length = float(r[dIdx['Длина']]) if 'Длина' in dIdx else 0
                    if width > 0 and height > 0 and length > 0:
                        volL = round(width * height * length / 1000, 3)
                    else:
                        volL = 0
                except Exception:
                    volL = 0
            dicts[wb_code] = {
                'subj': r[dIdx.get('Предмет', -1)] if 'Предмет' in dIdx else '',
                'volL': volL,
                'art_postav': r[dIdx.get('Артикул_поставщика', -1)] if 'Артикул_поставщика' in dIdx else '',
            }




        # 4. Комиссии
        print('📊 Чтение таблицы комиссий...')
        comm = {}
        for r in sh['comm'].range(2, 1).expand('table').value:
            subj = r[cIdx['Subject Name']]
            raw = str(r[cIdx['Commission, %']]).replace('%', '').replace(',', '.')
            if raw and raw != 'None':
                v = float(raw)
                comm[subj] = v / 100 if v > 1 else v

        # 5. Себестоимость
        print('📦 Чтение себестоимости...')
        cogs = {}
        for r in sh['cost'].range(2, 1).expand('table').value:
            key = f"{r[sIdx['Организация']]}|{r[sIdx['Артикул_поставщика']]}"
            cogs[key] = {
                'rub': round(to_num(r[sIdx['Себестоимость_руб']])),
                'rubWo': round(to_num(r[sIdx['Себестоимость_без_НДС_руб']]))
            }

        # 6. Параметры из "Настройки"
        print('⚙️ Чтение параметров из Настройки...')
        last_row = sh['cfg'].range('A' + str(sh['cfg'].cells.last_cell.row)).end('up').row
        cfg_raw = sh['cfg'].range(f"A2:B{last_row}").value if last_row >= 2 else []
        cfg = {k: to_num(v) for k, v in cfg_raw if k}
        T_FIRST = cfg.get('Логистика стоимость первого литра', 60)
        T_NEXT  = cfg.get('Логистика стоимость дополнительного литра', 16)
        T_COEF  = cfg.get('Коэффициент логистики', 115)
        STORE   = cfg.get('Хранение стоимость за шт.', 20)
        DRR     = cfg.get('ДРР', 15)

        # 7. Продажи и выручка
        print('📥 Загрузка планов продаж и выручки...')
        sales_data = sh['plan_sales'].range(2, 1).expand('table').value
        rev_data   = sh['plan_rev'].range(2, 1).expand('table').value

        # 8. Основной расчет
        print('🔄 Начинаем обработку строк...')
        out = []
        skipped = 0

        for rowIdx, ps in enumerate(sales_data):
            org = ps[pIdx['Организация']]
            wb_code = wb_code_key(ps[pIdx['Артикул_WB']])
                    # ↓ Вместо art = norm_key(...), теперь wb_code

            if not wb_code or str(org).lower().startswith('итого'):
                continue

            if wb_code not in dicts:
                print(f"[WARN] Артикул_WB не найден в справочнике: |{wb_code}|")
                print("Доступные артикула (фрагмент):", list(dicts.keys())[:10])

            meta  = dicts.get(wb_code, {'subj': '', 'volL': 0, 'art_postav': ''})
            print(f"[DEBUG] Предмет найден: {meta['subj']} для Артикул_WB {wb_code}")

            rate  = comm.get(meta['subj'], 0)
            cKey  = f"{org}|{meta['art_postav']}"  # можно оставить как раньше, если у тебя в себестоимости ключ через Артикул_поставщика
            unitC = cogs.get(cKey, {'rub': 0, 'rubWo': 0})
            if cKey not in cogs:
                skipped += 1

            pr = rev_data[rowIdx]
            for idx, mKey in enumerate(MONTHS):
                qty = to_num(ps[pIdx.get(mKey, -1)]) if mKey in pIdx else 0
                if not qty:
                    continue
                rev = round(to_num(pr[rIdx.get(mKey, -1)])) if mKey in rIdx else 0

                vol = meta['volL'] if meta['volL'] else 0
                if vol < 1:
                    perUnitLog = T_FIRST * T_COEF
                else:
                    extra_liters = math.ceil(vol - 1)
                    perUnitLog = (T_FIRST + extra_liters * T_NEXT) * T_COEF

                # --- ДОБАВЛЕНО: учёт % выкупа и обратной логистики ---
                REVERSE_LOG = 50  # обратная доставка за возврат единицы товара

                # Получаем процент выкупа для nmId (по справочнику)
                wb_percent = redemption_rate.get(str(wb_code), 95)  # 95% если нет данных
                return_rate = 1 - wb_percent / 100

                # Итоговая логистика на 1 ед. с учётом возвратов
                perUnitLog_full = perUnitLog + REVERSE_LOG * return_rate
                logiRub = round(perUnitLog_full * qty)


                commRub  = round(rev * rate)
                advRub   = round(rev * DRR)
                expMP    = commRub + logiRub + STORE * qty + advRub

                # ВЫВОДИМ В РЕЗУЛЬТАТ Артикул_WB как 2-ю колонку!
                out.append([
                    org, wb_code, meta['art_postav'], meta['subj'], str(idx + 1).zfill(2),
                    qty, rev, rate,
                    commRub, logiRub,
                    STORE * qty, advRub,
                    expMP,
                    unitC['rub']   * qty,
                    unitC['rubWo'] * qty
                ])


        print(f'✅ Обработка завершена. Всего строк для вывода: {len(out)}')
        if skipped:
            print(f'Skipped items due to missing COGS: {skipped}')

        # 9. Корректное формирование и запись умной таблицы
        hdr = [
        'Организация', 'Артикул_WB', 'Артикул_поставщика', 'Предмет', 'Месяц',
        'Кол-во, шт',  'Выручка, ₽', 'Комиссия WB %', 'Комиссия WB, ₽',
        'Логистика, ₽','Хранение, ₽','Реклама, ₽','Расходы МП, ₽',
        'СебестоимостьПродажРуб', 'СебестоимостьПродажБезНДС'
    ]


        def clean_number(x):
            if x is None or x == '':
                return 0
            try:
                return float(str(x).replace('₽', '').replace(' ', '').replace(',', '.'))
            except Exception:
                return 0

        rub_cols = [i for i, h in enumerate(hdr) if '₽' in h or 'Себестоимость' in h]
        pct_col = next((i for i, h in enumerate(hdr) if 'Комиссия WB %' in h), None)

        cleaned_out = []
        for row in out:
            cleaned_row = []
            for i, v in enumerate(row):
                if i in rub_cols:
                    cleaned_row.append(clean_number(v))
                elif pct_col is not None and i == pct_col:
                    cleaned_row.append(clean_number(v))
                else:
                    cleaned_row.append(v)
            cleaned_out.append(cleaned_row)

        res = sh['result']
        res.clear()
        header_rng = res.range((1, 1), (1, len(hdr)))
        header_rng.value = hdr

        if cleaned_out:
            data_rng = res.range((2, 1), (len(cleaned_out) + 1, len(hdr)))
            data_rng.value = cleaned_out

        try:
            for tbl in res.api.ListObjects:
                if tbl.Name == 'РасходыТаблица':
                    tbl.Delete()
        except Exception:
            pass

        last_row = len(cleaned_out) + 1
        last_col = len(hdr)
        table_range = res.range((1, 1), (last_row, last_col))
        lo = res.api.ListObjects.Add(
            SourceType=1,
            Source=table_range.api,
            XlListObjectHasHeaders=1
        )
        lo.Name = 'РасходыТаблица'
        lo.TableStyle = 'TableStyleMedium7'

        for idx in rub_cols:
            c = idx + 1
            res.range((2, c), (last_row, c)).api.NumberFormat = '0 ₽'
        if pct_col is not None:
            c = pct_col + 1
            res.range((2, c), (last_row, c)).api.NumberFormat = '0%'

        total_row = last_row + 1
        res.range((total_row, 1)).value = 'ИТОГО'
        res.range((total_row, 1)).api.Font.Bold = True
        for c in [5, 6, 8, 9, 10, 11, 12, 13, 14]:
            col = col_letter(c)
            res.range((total_row, c)).formula = f'=SUBTOTAL(9,{col}2:{col}{total_row-1})'
            res.range((total_row, c)).api.Font.Bold = True
            res.range((total_row, c)).api.NumberFormat = '0 ₽'

        res.autofit()
        res.api.Rows.AutoFit()

        print('🟩 Итоговая таблица создана, форматирование применено.')
        wb.save()
        print(f'🏁 [FINISH] Выполнение завершено за {round(time.time() - start, 1)} сек.')
        sheet = sh['result']
        sheet.api.Tab.Color = 9687200  # насыщенно-зелёный

        # Если хочешь лист последним:
        # wb.sheets.move(sheet, after=wb.sheets[-1])

        # Если хочешь лист вторым:
        sheet.api.Move(Before=wb.sheets[9].api)


        print('✅ Цвет вкладки и порядок листов применены.')

    finally:
        # Закрываем только если не из Excel
        if not from_caller and wb is not None:
            wb.close()
        if app is not None:
            app.quit()
            del app

if __name__ == '__main__':
    main()

    