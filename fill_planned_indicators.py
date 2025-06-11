# fill_planned_indicators.py
# -------------------------------------------------------------------
# Пересчёт плановых показателей и налогов  (v1.7 — 05-06-2025)
# -------------------------------------------------------------------
# • Лист «РасчетПлановыхПоказателей» = 3-й, ярлык зелёный
# • Умная таблица PlannedIndicatorsTbl, стиль TableStyleMedium7
# • Строка TotalsRow: подпись «Итого» + суммы
# • Все рублевые колонки → формат "#,##0 ₽"
# • Оптимизированы COM-вызовы: экран/события/калькуляция Off во
#   время тяжёлых операций — «виснуть» больше не будет
# -------------------------------------------------------------------

import os
import argparse
import xlwings as xw

# ---------- 1. CLI --------------------------------------------------------
def parse_args():
    p = argparse.ArgumentParser(add_help=False,
                                description='Пересчёт плановых показателей')
    p.add_argument('-f', '--file', default='Finmodel.xlsm',
                   help='Имя Excel-книги (по умолчанию Finmodel.xlsm)')
    args, _ = p.parse_known_args()       # игнорируем лишние флаги xlwings
    return args
ARGS = parse_args()

# ---------- 2. Пути и имена листов ----------------------------------------
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, ARGS.file)

SHEET_WB   = 'РасчётЭкономикиWB'
SHEET_OZON = 'РасчетЭкономикиОзон'
SHEET_ORG  = 'НастройкиОрганизаций'
SHEET_SAL  = 'Зарплата'
SHEET_OTH  = 'ПрочиеРасходы'
SHEET_OUT  = 'РасчетПлановыхПоказателей'

TABLE_NAME  = 'PlannedIndicatorsTbl'
TABLE_STYLE = 'TableStyleMedium7'          # зелёный Medium 7

LIMIT_GROSS_USN = 450_000_000              # ₽

# ---------- 3. Вспомогательные функции ------------------------------------
def open_wb():
    """Возвращает (wb, app). app == None, если скрипт вызван из Excel."""
    try:                    # вызов из макроса RunPython
        wb, app = xw.Book.caller(), None
        print(f'→ Excel-режим: {wb.fullname}')
    except Exception:       # запуск из терминала
        app = xw.App(visible=False, add_book=False)
        wb  = app.books.open(EXCEL_PATH, read_only=False)
        print(f'→ Консоль-режим: {EXCEL_PATH}')
    return wb, app

parse_money = lambda v: float(
    ''.join(c for c in str(v).replace(' ', '').replace('₽', '')
            if c.isdigit() or c in '-.')) if v not in (None, '') else 0.0

def parse_month(val):
    """
    Корректно извлекает месяц из числа, float, строки или формата '01.2024'
    """
    # Excel часто читает числа как float (1.0, 2.0, ...)
    if isinstance(val, float) and val.is_integer():
        return int(val)
    if isinstance(val, int):
        return val
    s = str(val).strip()
    # поддержка строки "10", "01.2024", "2024-03"
    if s.isdigit():
        return int(s)
    # Если формат "01.2024"
    if '.' in s:
        s = s.split('.')[0]
    elif '-' in s:
        s = s.split('-')[-1]
    if s.isdigit():
        return int(s)
    return 0

def log_month(*args, **kwargs):
    pass
def build_idx(header): return {str(c).strip().lower(): i for i, c in enumerate(header)}

def read_rows(sh):
    rng = sh.range(1, 1).expand('table').value
    return (rng[1:], build_idx(rng[0])) if rng and len(rng) > 1 else ([], {})

def acc(iterable, kfn, vfn):
    d = {}
    for x in iterable:
        d[kfn(x)] = d.get(kfn(x), 0) + vfn(x)
    return d

def ndfl_prog(base):
    left, tax, prev = base, 0, 0
    for lim, r in [(2.4e6, .13), (5e6, .15),
                   (20e6, .18), (50e6, .20), (float('inf'), .22)]:
        take = min(left, lim - prev)
        tax += take * r
        left -= take
        prev = lim
        if left <= 0:
            break
    return tax

def nds_rate(prev, curr, mode, def_r):
    if mode == 'ОСНО' or curr > 450e6:
        return 20
    if prev > 250e6:
        return max(7, def_r)
    if prev > 60e6:
        return max(5, def_r)
    return def_r

def log_nds(month, org, prev, curr, mode, rate, lvl):
    """Короткий лог выбора ставки НДС.
       lvl = 'M' — для общей (Monthly), 'O' — для организации."""
    print(f"[NDS-{lvl}] {month:>2} | {org:<20} | prev={prev:,.0f} "
          f"→ curr={curr:,.0f} | mode={mode:<8} | rate={rate}%")



# ---------- 4. Главная функция --------------------------------------------
def fill_planned_indicators():
    headers = ['Организация', 'Месяц', 'Выручка, ₽', 'Выручка накоп., ₽',
           'Выручка сводно, ₽', 'Выручка без НДС, ₽', 'НДС, ₽',
           'Ставка НДС, %', 'Себестоимость руб', 'Себестоимость без НДС',
           'Расх. MP с НДС, ₽',          # ← новая колонка (брутто)
           'Расх. MP без НДС, ₽',        # ← бывшая «Расх. MP, ₽»
           'ФОТ, ₽', 'ЕСН, ₽', 'Прочие, ₽', 'EBITDA, ₽',
           'EBITDA накоп., ₽', 'EBITDA сводно, ₽', 'Режим',
           'Ставка УСН, %', 'Налог, ₽', 'Чистая прибыль, ₽']

    ruble_cols = [h for h in headers if '₽' in h or h.startswith('Себестоимость')]

    wb = app = None
    try:
        # === 4.1 Открываем книгу ========================================
        wb, app = open_wb()
        ss = wb
        sheet_names = [s.name for s in ss.sheets]

        # === 4.2 Данные WB =============================================
        # === 4.2 Данные WB =============================================
        if SHEET_WB not in sheet_names:
            raise ValueError(f'Нет листа {SHEET_WB}')

        # ❶ читаем строки и индексы
        wb_rows, wb_idx = read_rows(ss.sheets[SHEET_WB])

        # выводим индексы только при запуске из ТЕРМИНАЛА (app == None)
        if app is None:          # <<< добавили условие
            print('WB idx:', wb_idx)

        # ❸ проверяем обязательные колонки
        need_wb = [
            'организация', 'месяц', 'выручка, ₽', 'расходы мп, ₽',
            'себестоимостьпродажруб', 'себестоимостьпродажбезндс'
        ]
        for col in need_wb:
            if col not in wb_idx:
                raise ValueError(f'Колонка «{col}» отсутствует в {SHEET_WB}')


        # === 4.3 Данные Ozon ===========================================

        rows = []
        oz_rows = []                      # на случай отсутствия листа Ozon
                         # сюда будем складывать все строки
        if SHEET_OZON in sheet_names:
            oz_rows, oz_idx_raw = read_rows(ss.sheets[SHEET_OZON])

            # Приводим ключи к нижнему регистру и убираем пробелы
            oz_idx = {str(k).strip().lower(): i for k, i in oz_idx_raw.items()}

            if app is None:          # <<< только из терминала
                print('Ozon idx:', oz_idx)


            need_oz = [
                'организация', 'месяц', 'выручка_руб', 'итогорасходымп_руб',
                'себестоимостьпродаж_руб', 'себестоимостьбезндс_руб'
            ]
            for col in need_oz:
                if col not in oz_idx:
                    raise ValueError(f'Колонка «{col}» отсутствует в {SHEET_OZON}')


        for i, r in enumerate(oz_rows, 2):  # 2 — потому что range(1,1) = A1, а данные с 2-й строки
            org = r[oz_idx['организация']]
            raw_month = r[oz_idx['месяц']]
            if not org or str(org).strip().lower() in ('итого', 'total'):
                continue
            month = parse_month(raw_month)
            if month == 0 or not (1 <= month <= 12):
                log_month(raw_month, src='Ozon', rownum=i, reason=f'игнорируется, результат parse_month={month}')
                continue
            log_month(raw_month, src='Ozon', rownum=i, reason=f'принят, результат parse_month={month}')
            rows.append(dict(
                org=org,
                month=month,
                rev=parse_money(r[oz_idx['выручка_руб']]),
                mp=parse_money(r[oz_idx['итогорасходымп_руб']]),
                cr=parse_money(r[oz_idx['себестоимостьпродаж_руб']]),
                cn=parse_money(r[oz_idx['себестоимостьбезндс_руб']])
            ))


        # === 4.4 Добавляем строки WB ====================================
        for i, r in enumerate(wb_rows, 2):
            org = r[wb_idx['организация']]
            raw_month = r[wb_idx['месяц']]
            if not org or str(org).strip().lower() in ('итого', 'total'):
                continue
            month = parse_month(raw_month)
            if month == 0 or not (1 <= month <= 12):
                log_month(raw_month, src='WB', rownum=i, reason=f'игнорируется, результат parse_month={month}')
                continue
            log_month(raw_month, src='WB', rownum=i, reason=f'принят, результат parse_month={month}')
            rows.append(dict(
                org=org, month=month,
                rev=parse_money(r[wb_idx['выручка, ₽']]),
                mp=parse_money(r[wb_idx['расходы мп, ₽']]),
                cr=parse_money(r[wb_idx['себестоимостьпродажруб']]),
                cn=parse_money(r[wb_idx['себестоимостьпродажбезндс']])
            ))

        if not rows:
            print('⚠️  Нет данных — выходим'); return

        # === 4.5 НастройкиОрганизаций ===================================
        if SHEET_ORG not in sheet_names:
            raise ValueError(f'Нет листа {SHEET_ORG}')
        cfg_rows, cfg_idx = read_rows(ss.sheets[SHEET_ORG])
        org_cfg = {}
        for r in cfg_rows:
            org = r[cfg_idx['организация']]

            # --- ставка НДС/УСН (как было) ---
            nds = parse_money(str(r[cfg_idx.get('ставка ндс', '')]).replace('%', '').replace(',', '.'))
            nds = nds * 100 if 0 < nds < 1 else nds
            usn = parse_money(str(r[cfg_idx.get('ставканалогаусн', '')]).replace('%', '').replace(',', '.'))
            usn = usn * 100 if 0 < usn < 1 else usn

            # --- режим налогообложения ---
            col_new = cfg_idx.get('режимналогооблnew')
            col_old = cfg_idx.get('режим_налогообложения')     # оставим поддержку старого
            mode_val = 'ОСНО'                                   # дефолт
            src_col  = 'default'
            if col_new is not None and str(r[col_new]).strip():
                mode_val = str(r[col_new]).strip()
                src_col  = 'New'
            elif col_old is not None and str(r[col_old]).strip():
                mode_val = str(r[col_old]).strip()
                src_col  = 'Old'

            # логируем выбор
            if app is None:
                print(f"[CFG] {org:<20} режим ← {src_col}: {mode_val}")

            org_cfg[org] = dict(
                type=str(r[cfg_idx.get('тип_организации', '')]).strip() or 'ООО',
                orig_mode=mode_val,
                consolidation=str(r[cfg_idx.get('консолидация', '')]).strip().lower() != 'нет',
                nds_rate=nds,
                usn_rate=usn
            )


        # === 4.6 Зарплата и прочие расходы ==============================
        salary = {}
        if SHEET_SAL in sheet_names:
            sal_rows, sal_idx = read_rows(ss.sheets[SHEET_SAL])
            for r in sal_rows:
                salary[r[sal_idx['организация']]] = dict(
                    fot=parse_money(r[sal_idx['фот']]),
                    mode=str(r[sal_idx['режим_зп']]).strip())

        other = {}
        if SHEET_OTH in sheet_names:
            oth_rows, _ = read_rows(ss.sheets[SHEET_OTH])
            for org, val in oth_rows:
                other[org] = parse_money(val)

        # === 4.7 Группировка по (org, month) ============================
        grouped = {}
        for r in rows:
            k = (r['org'], r['month'])
            g = grouped.setdefault(k, dict(org=r['org'], month=r['month'],
                                            rev=0, mp=0, cr=0, cn=0))
            for f in ('rev', 'mp', 'cr', 'cn'):
                g[f] += r[f]

        records = sorted(grouped.values(), key=lambda x: x['month'])

        rev_m = acc(records, lambda x: x['month'], lambda x: x['rev'])
        months = sorted(rev_m)
        cum_all, s = {}, 0
        for m in months:
            s += rev_m[m]; cum_all[m] = s

# --- ставка НДС по консолидированному обороту на каждый месяц ---

        nds_by_month = {}
        prev_gross = 0
        for m in months:
            curr_gross = cum_all[m]
            # здесь ДОЛЖНО быть 'Доходы', а не 'ОСНО'
            rateM = nds_rate(prev_gross, curr_gross, 'Доходы', 0)
            nds_by_month[m] = rateM
            log_nds(m, 'ALL', prev_gross, curr_gross, 'CONS', rateM, 'M')
            prev_gross = curr_gross




        cum_org = {org: {} for org in {g['org'] for g in records}}
        for org in cum_org:
            run = 0
            for m in months:
                run += sum(g['rev'] for g in records if g['org'] == org and g['month'] == m)
                cum_org[org][m] = run

        # === 4.8 Основной расчёт =======================================

        p_rev, p_ebit, p_net, last_mode = {}, {}, {}, {}
        out = []

        for g in records:
            cfg = org_cfg.get(g['org'], dict(orig_mode='ОСНО', consolidation=False,
                                            nds_rate=0, usn_rate=0, type='ООО'))
            mode_eff = cfg['orig_mode']

            # --- выбор “гросс” выручки ---
            gross = cum_all[g['month']] if cfg['consolidation'] else cum_org[g['org']][g['month']]
            if cfg['orig_mode'] in ('Доходы', 'Доходы-Расходы') and gross > LIMIT_GROSS_USN:
                mode_eff = 'ОСНО'

            
    # --- ставка НДС ---
            if cfg['consolidation']:
                nds = nds_by_month[g['month']]            # ❶ сначала значение
                prev_g = cum_all.get(g['month'] - 1, 0)
                curr_g = cum_all[g['month']]
                log_nds(g['month'], g['org'], prev_g, curr_g, mode_eff, nds, 'O')  # ❷ потом лог
            else:
                prev = p_rev.get(g['org'], 0)
                nds  = nds_rate(prev, prev + g['rev'], mode_eff, cfg['nds_rate'])   # ❶
                prev_g, curr_g = prev, prev + g['rev'] 

        # --- нижний предел из «Ставка НДС» в настройках ---
            nds = max(nds, cfg['nds_rate'])         # ← ДОБАВЛЕННАЯ строка

            # --- лог после окончательного значения ---
            log_nds(g['month'], g['org'], prev_g, curr_g, mode_eff, nds, 'O')
            # ---------- расчёт показателей ----------
            revN    = g['rev'] / (1 + nds / 100)
            nds_sum = g['rev'] - revN

            mpGross = g['mp']
            mpNet   = mpGross / 1.2

            sal = salary.get(g['org'], dict(fot=0, mode='Неформ'))
            fot = sal['fot']
            esn = fot * 0.30 if sal['mode'] == 'Официальная' else 0
            oth_cost = other.get(g['org'], 0)

            cost_base = g['cn'] if round(nds) == 20 else g['cr']
            ebit = revN - (cost_base + mpNet + fot + esn + oth_cost)

            # --- аккумулируем ---
            p_rev[g['org']]  = p_rev.get(g['org'], 0) + g['rev']
            p_ebit[g['org']] = p_ebit.get(g['org'], 0) + ebit
            p_net[g['org']]  = p_net.get(g['org'], 0) + revN

            out.append(dict(
                org=g['org'], m=g['month'], rev=g['rev'], cumG=gross,
                revN=revN, ndsSum=nds_sum, nds=nds,
                cr=g['cr'], cn=g['cn'],
                mpGross=mpGross, mpNet=mpNet,
                fot=fot, esn=esn, oth=oth_cost, ebit=ebit,
                cumN=p_net[g['org']], cumE=p_ebit[g['org']],
                mode=mode_eff, type=cfg['type'], prevM=last_mode.get(g['org']),
                usn=cfg['usn_rate'])
            )
            last_mode[g['org']] = mode_eff


        ebit_m = acc(out, lambda x: x['m'], lambda x: x['ebit'])
        rows_out, cum_osno = [], {}
        for r in out:
            tax = base = 0
            rate = '0%'
            if r['mode'] == 'Доходы':
                base = max(r['revN'], 0); tax = round(base * r['usn'] / 100); rate = f"{r['usn']}%"
            elif r['mode'] == 'Доходы-Расходы':
                base = max(r['ebit'], 0); tax = round(base * r['usn'] / 100); rate = f"{r['usn']}%"
            else:  # ОСНО
                if r['type'] == 'ИП':
                    if r['prevM'] != 'ОСНО': cum_osno[r['org']] = 0
                    base = max(r['ebit'], 0)
                    prev = cum_osno.get(r['org'], 0); cum = prev + base
                    tax = round(ndfl_prog(cum) - ndfl_prog(prev)); cum_osno[r['org']] = cum
                    rate = f"{(tax / base * 100):.2f}%" if base else '0%'
                else:
                    base = max(r['ebit'], 0); tax = round(base * 0.25); rate = '25%'
            rows_out.append([
                #  1  Организация
                r['org'],
                #  2  Месяц
                r['m'],
                #  3  Выручка, ₽
                round(r['rev']),
                #  4  Выручка накоп., ₽
                round(r['cumG']),
                #  5  Выручка сводно, ₽
                round(cum_all[r['m']]),
                #  6  Выручка без НДС, ₽
                round(r['revN']),
                #  7  НДС, ₽
                round(r['ndsSum']),
                #  8  Ставка НДС, %
                f"{round(r['nds'])}%",
                #  9  Себестоимость руб
                round(r['cr']),
                # 10  Себестоимость без НДС
                round(r['cn']),
                # 11  Расх. MP с НДС, ₽   (брутто)
                round(r['mpGross']),
                # 12  Расх. MP без НДС, ₽ (нетто)
                round(r['mpNet']),
                # 13  ФОТ, ₽
                round(r['fot']),
                # 14  ЕСН, ₽
                round(r['esn']),
                # 15  Прочие, ₽
                round(r['oth']),
                # 16  EBITDA, ₽
                round(r['ebit']),
                # 17  EBITDA накоп., ₽
                round(r['cumE']),
                # 18  EBITDA сводно, ₽
                round(ebit_m[r['m']]),
                # 19  Режим
                r['mode'],
                # 20  Ставка УСН, %
                rate,
                # 21  Налог, ₽
                tax,
                # 22  Чистая прибыль, ₽
                round(r['ebit'] - tax)
            ])


        # === 4.9 Запись в Excel ====================================
        
        target = None
        for sht in ss.sheets:
                clean = sht.name.replace('\u200b', '').strip()   # убираем нулевой-ширины пробелы
                if clean == SHEET_OUT:
                    target = sht
                    break

        if target is None:                       # листа нет — создаём
                target = ss.sheets.add(SHEET_OUT)

        sh = target
        sh.clear()                          # очищаем старые данные

        sh.range(1, 1).value = headers
        if rows_out:
            sh.range(2, 1).value = rows_out

        # ------ создаём / обновляем умную таблицу (оптимизировано) ------
       
        screen, calc = wb.app.screen_updating, wb.app.calculation
        events       = wb.app.enable_events
        wb.app.screen_updating = False
        wb.app.enable_events   = False
        wb.app.calculation     = 'manual'

        try:
            # 1) диапазон данных
            last_row = sh.range(1, 1).end('down').row
            last_col = sh.range(1, 1).end('right').column
            lo_range = sh.range((1, 1), (last_row, last_col)).api

            # 2) удалить старую PlannedIndicatorsTbl, если была
            for lo in list(sh.api.ListObjects):
                if lo.Name == TABLE_NAME:
                    lo.Delete()

            # 3) создать новую ListObject без TotalsRow
            lo = sh.api.ListObjects.Add(1, lo_range, None, 1)
            lo.Name, lo.TableStyle = TABLE_NAME, TABLE_STYLE   # стиль Medium 7
            fmt_fin = (
                '_-* #,##0 ₽_-;'           # положительные
                '_-* (#,##0 ₽)_-;'         # отрицательные (скобки)
                '_-* "-"?? ₽_-;'           # нули → тире
                '_-@_-'                    # текст
            )


            # 4) форматируем все ₽-колонки единым вызовом
            fmt = fmt_fin
            ruble_idx = [headers.index(c) + 1 for c in ruble_cols]
            for i in ruble_idx:
                lo.ListColumns(i).Range.NumberFormat = fmt

        finally:
            wb.app.calculation     = calc
            wb.app.enable_events   = events
            wb.app.screen_updating = screen

        # ------ «псевдо-итого» сразу под таблицей -----------------------
        from xlwings.utils import col_name

        total_row = last_row + 1                 # ← исправили
        sh.range(total_row, 1).value = 'Итого'

        for idx, col in enumerate(headers, start=1):
            if col in ruble_cols:
                letter = col_name(idx)
                sh.range(total_row, idx).formula = \
                    f"=SUBTOTAL(109,{letter}$2:{letter}${last_row})"
                sh.range(total_row, idx).number_format = fmt


        # ------ ярлык и позиция листа ----------------------------------
        sh.api.Tab.ColorIndex = 10
        if sh.index != 3:
            sh.api.Move(Before=ss.sheets[8].api)

        print(f'✔️  Готово! Записано строк: {len(rows_out)}')

    finally:
        if wb:
            wb.save()
            if app:
                wb.close(); app.quit()

# ---------- 5. Точка входа -----------------------------------------------

if __name__ == '__main__':
    fill_planned_indicators()

