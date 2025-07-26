# fill_planned_indicators.py
# -------------------------------------------------------------------
# Пересчёт плановых показателей и налогов  (v1.8 — 26-07-2025)
# -------------------------------------------------------------------
# • Лист «РасчетПлановыхПоказателей» = 3-й, ярлык зелёный
# • Умная таблица PlannedIndicatorsTbl, стиль TableStyleMedium7
# • Строка TotalsRow: подпись «Итого» + суммы
# • Все рублевые колонки → финансовый формат (красный минус)
# • Оптимизированы COM-вызовы: экран/события/калькуляция Off во
#   время тяжёлых операций — «виснуть» больше не будет
# -------------------------------------------------------------------
# CHANGELOG
# v1.8 — 26-07-2025: фикс расчёта налога при отрицательной
#                     консолидированной базе после перехода на ОСНО

import os
import argparse
import xlwings as xw
import logging

# ---------- Логирование в файл -------------------------------------
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'log')
os.makedirs(LOG_DIR, exist_ok=True)
LOG_PATH = os.path.join(LOG_DIR, 'fill_planned_indicators.log')

logging.basicConfig(
    filename=LOG_PATH,
    filemode='w',
    level=logging.INFO,
    format='[%(asctime)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
def log_info(msg):
    print(msg)               # ✔️ выводит в терминал
    logging.info(msg)        # ✔️ записывает в log/fill_planned_indicators.log

# По умолчанию подавляем сообщения о месяце. Включается CLI-флагом
# ``-dm``/``--debug-month``.
DEBUG_MONTH = False


# ---------- 1. CLI --------------------------------------------------------
def parse_args():
    p = argparse.ArgumentParser(add_help=False,
                                description='Пересчёт плановых показателей')
    p.add_argument('-f', '--file', default='excel/Finmodel.xlsm',
                   help='Имя Excel-книги (по умолчанию excel/Finmodel.xlsm)')
    p.add_argument('-dm', '--debug-month', action='store_true',
                   help='log every imported month')
    args, _ = p.parse_known_args()       # игнорируем лишние флаги xlwings
    global DEBUG_MONTH
    DEBUG_MONTH = args.debug_month
    return args
ARGS = parse_args()

# ---------- 2. Пути и имена листов ----------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
EXCEL_PATH = os.path.join(PROJECT_DIR, ARGS.file)


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
        log_info(f'→ Excel-режим: {wb.fullname}')
    except Exception:       # запуск из терминала
        app = xw.App(visible=False, add_book=False)
        wb  = app.books.open(EXCEL_PATH, read_only=False)
        log_info(f'→ Консоль-режим: {EXCEL_PATH}')
    return wb, app

def parse_money(v):
    if v in (None, ''):
        return 0.0
    s = str(v).replace(' ', '').replace('₽', '').replace(',', '.')
    s = ''.join(
        c
        for c in s
        if c.isdigit() or c in '-.'
    )
    return float(s)

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
    """Log month parsing results during data import."""
    if not DEBUG_MONTH:      # ➜ молчим, если не debug-режим
        return
    if not args and not kwargs:
        return
    val = args[0] if args else kwargs.get('val') or kwargs.get('value')
    src = kwargs.get('src', '')
    rownum = kwargs.get('rownum', '')
    reason = kwargs.get('reason', '')
    msg = f"[MONTH] src={src} row={rownum} value={val!r} {reason}"
    log_info(msg)
def build_idx(header): return {str(c).strip().lower(): i for i, c in enumerate(header)}

def read_rows(sh):
    rng = sh.range(1, 1).expand('table').value
    return (rng[1:], build_idx(rng[0])) if rng and len(rng) > 1 else ([], {})

def acc(iterable, kfn, vfn):
    d = {}
    for x in iterable:
        d[kfn(x)] = d.get(kfn(x), 0) + vfn(x)
    return d

def find_key(idx, target):
    """Return key from ``idx`` that matches ``target`` ignoring spaces,
    underscores and punctuation.

    Both ``idx`` keys and ``target`` may contain characters like commas or the
    currency sign ``₽`` which should be ignored when searching.  The function
    returns the original key from ``idx`` if a normalized match is found.
    """

    def norm(text):
        return ''.join(c for c in str(text).lower() if c.isalnum())

    target_norm = norm(target)
    for k in idx:
        if norm(k) == target_norm:
            return k
    return None

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
    msg = f"[NDS-{lvl}] {month:>2} | {org:<20} | prev={prev:,.0f} → curr={curr:,.0f} | mode={mode:<8} | rate={rate}%"
    log_info(msg)


def full_cogs(cn, nds):
    """Return cost including non-refundable VAT for reduced rates."""
    return cn * (1 + nds / 100)


def _calc_row(
    revN,
    mpNet,
    cost_sales,
    cost_tax,
    fot,
    esn,
    oth,
    mode,
    mpGross=0,
    oklad_of=None,
):
    """Calculate management and tax EBITDA for given inputs."""

    if oklad_of is None:
        oklad_of = fot

    ebit_mgmt = revN - (cost_sales + mpNet + fot + esn + oth)
    if mode == 'Доходы-Расходы':
        ebit_tax = revN - (cost_tax + mpGross + oklad_of + esn + oth)
    else:
        ebit_tax = ebit_mgmt
    return {
        'EBITDA, ₽': ebit_mgmt,
        'Расчет_базы_налога': ebit_tax,
    }


def _apply_consolidated_dr_tax(rows):
    """Distribute USN "Доходы-Расходы" tax within consolidated rows."""
    from collections import defaultdict

    grouped = defaultdict(list)
    for r in rows:
        grouped[r['m']].append(r)

    totals = {}
    for m, items in grouped.items():
        total_rev = sum(x['revN'] for x in items)
        total_ebit = sum(x['ebit_tax'] for x in items)
        rate = items[0]['usn'] / 100
        real_tax = rate * total_ebit
        min_tax = total_rev * 0.01
        tax_sum = max(real_tax, min_tax)
        forced_min = real_tax < min_tax
        for r in items:
            share = r['revN'] / total_rev if total_rev else 0
            r['tax'] = round(tax_sum * share)
            r['usn_forced_min'] = forced_min
        totals[m] = total_ebit
    return totals


def calc_consolidated_min_tax(base, revenue, rate):
    """Return USN tax with 1% minimum for consolidated mode."""
    real_tax = base * rate
    min_tax = revenue * 0.01
    return max(real_tax, min_tax)


def consolidate_osno_tax(rows, meta):
    """
    В режиме консолидации ОСНО/ИП:
    - налог только в одной строке на месяц (остальные — ноль)
    - в остальных строках Чистая прибыль = EBITDA
    """
    from collections import defaultdict

    grouped = defaultdict(list)
    for i, m in enumerate(meta):
        if m.get('consolidation') and m.get('type') == 'ИП' and m.get('mode') == 'ОСНО':
            grouped[(m['m'])].append(i)  # группируем по месяцу

    for m, idxs in grouped.items():
        if not idxs:
            continue
        # Найти первую строку по алфавиту организации — туда записываем налог
        main_idx = sorted(idxs, key=lambda i: rows[i][0])[0]
        # Суммируем налог по всем строкам месяца
        total_tax = sum(rows[i][28] for i in idxs)
        # В остальные строки — налог = 0, чистая прибыль = EBITDA
        for i in idxs:
            if i == main_idx:
                rows[i][28] = total_tax
                rows[i][29] = rows[i][19] - total_tax  # Чистая прибыль = EBITDA - налог
            else:
                rows[i][28] = 0
                rows[i][29] = rows[i][19]  # Чистая прибыль = EBITDA





# ---------- 4. Главная функция --------------------------------------------
def fill_planned_indicators():
    headers = [
        'Организация', 'Месяц', 'Выручка, ₽', 'Выручка накоп., ₽',
        'Выручка сводно, ₽', 'Выручка без НДС, ₽', 'НДС, ₽',
        'Ставка НДС, %', 'Себестоимость руб', 'Себестоимость без НДС',
        'Себестоимость Налог, ₽', 'Себестоимость Налог без НДС, ₽',
        'Расх. MP с НДС, ₽',          # ← новая колонка (брутто)
        'Расх. MP без НДС, ₽',        # ← бывшая «Расх. MP, ₽»
        'ФОТ, ₽', 'Оклад_Оф, ₽', 'ЕСН, ₽', 'Прочие, ₽', 'EBITDA, ₽',
        'Расчет_базы_налога', 'EBITDA нал. накоп., ₽',
        'EBITDA накоп., ₽', 'EBITDA сводно, ₽', 'РасчетБазыНалогаНакопКонсол',
        'БазаНДФЛ ОСНО накоп., ₽', 'БазаНДФЛ ОСНО накоп. сводно, ₽', 'Режим',
        'Ставка УСН, %', 'Налог, ₽', 'Чистая прибыль, ₽',
    ]

    ruble_cols = [h for h in headers if '₽' in h or h.startswith('Себестоимость')]
    for col in ['БазаНДФЛ ОСНО накоп., ₽', 'БазаНДФЛ ОСНО накоп. сводно, ₽']:
        if col not in ruble_cols:
            ruble_cols.append(col)

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
            log_info(f'WB idx: {wb_idx}')


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
                log_info(f'Ozon idx: {oz_idx}')



            need_oz = [
                'организация', 'месяц', 'выручка_руб', 'итогорасходымп_руб',
                'себестоимостьпродаж_руб', 'себестоимостьбезндс_руб'
            ]
            # Колонки себестоимости по налоговому учёту могут отсутствовать
            tax_col_candidates = [
                'СебестоимостьПродажНалог, ₽',
                'СебестоимостьНалог_руб',
            ]
            tax_col_oz = None
            for cand in tax_col_candidates:
                key = find_key(oz_idx, cand)
                if key is not None:
                    tax_col_oz = oz_idx[key]
                    break

            tax_nds_col_oz = None
            key = find_key(oz_idx, 'СебестоимостьПродажНалог_без_НДС, ₽')
            if key is not None:
                tax_nds_col_oz = oz_idx[key]

            _has_tax_cogs = tax_col_oz is not None
            _has_tax_cogs_wo = tax_nds_col_oz is not None
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
                cn=parse_money(r[oz_idx['себестоимостьбезндс_руб']]),
                ct=parse_money(r[tax_col_oz]) if tax_col_oz is not None else 0,
                ct_wo=parse_money(r[tax_nds_col_oz]) if tax_nds_col_oz is not None else 0
            ))


        # === 4.4 Добавляем строки WB ====================================
        tax_col_wb_key = find_key(wb_idx, 'СебестоимостьНалог') or \
                         find_key(wb_idx, 'СебестоимостьПродажНалог')
        tax_col_wb = wb_idx[tax_col_wb_key] if tax_col_wb_key is not None else None
        _has_tax_cogs_wb = tax_col_wb is not None

        tax_wo_col_wb_key = (find_key(wb_idx, 'СебестоимостьНалогБезНДС') or
                              find_key(wb_idx, 'СебестоимостьПродажНалогБезНДС'))
        tax_wo_col_wb = (wb_idx[tax_wo_col_wb_key]
                          if tax_wo_col_wb_key is not None else None)
        _has_tax_cogs_wo_wb = tax_wo_col_wb is not None
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
                cn=parse_money(r[wb_idx['себестоимостьпродажбезндс']]),
                ct=parse_money(r[tax_col_wb]) if tax_col_wb is not None else 0,
                ct_wo=parse_money(r[tax_wo_col_wb]) if tax_wo_col_wb is not None else 0
            ))

        if not rows:
            log_info('⚠️  Нет данных — выходим')
            return

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
                log_info(f"[CFG] {org:<20} режим ← {src_col}: {mode_val}")

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
            oth_rows, oth_idx_raw = read_rows(ss.sheets[SHEET_OTH])
            oth_idx = {str(k).strip().lower(): i for k, i in oth_idx_raw.items()}

            # проверим, что нужные колонки существуют
            for col in ('организация', 'расходы'):
                if col not in oth_idx:
                    raise ValueError(f'Колонка «{col}» отсутствует в {SHEET_OTH}')

            # Суммируем по каждой организации все "Расходы"
            for r in oth_rows:
                org = r[oth_idx['организация']]
                val = parse_money(r[oth_idx['расходы']])
                if org not in other:
                    other[org] = 0
                other[org] += val

        # --- 4.6A Суммарные значения ФОТ и ЕСН по организации ---
        SHEET_PAYROLL = 'РасчетЗарплаты'
        payroll_rows, payroll_idx = read_rows(ss.sheets[SHEET_PAYROLL])
        esn_by_org = {}
        fot_by_org = {}
        oklad_by_org = {}

        for r in payroll_rows:
            try:
                scenario = str(r[payroll_idx['сценарий']]).strip().lower()
                org = r[payroll_idx['организация']]
                if scenario == 'как есть' and org:
                    esn = float(r[payroll_idx['итого_взносы']] or 0)
                    fot = float(r[payroll_idx['итого_зарплата']] or 0)
                    ok_off = 0
                    if 'оклад_оф' in payroll_idx:
                        ok_off = float(r[payroll_idx['оклад_оф']] or 0)
                    esn_by_org[org] = esn_by_org.get(org, 0) + esn
                    fot_by_org[org] = fot_by_org.get(org, 0) + fot
                    oklad_by_org[org] = oklad_by_org.get(org, 0) + ok_off
            except Exception:
                pass

     


        # === 4.7 Группировка по (org, month) ============================
        grouped = {}
        for r in rows:
            k = (r['org'], r['month'])
            g = grouped.setdefault(
                k,
                dict(org=r['org'], month=r['month'], rev=0, mp=0, cr=0, cn=0, ct=0, ct_wo=0)
            )
            for f in ('rev', 'mp', 'cr', 'cn', 'ct', 'ct_wo'):
                g[f] += r.get(f, 0)

        records = sorted(grouped.values(), key=lambda x: x['month'])

        rev_m = acc(records, lambda x: x['month'], lambda x: x['rev'])
        months = sorted(rev_m)
        cum_all, s = {}, 0
        for m in months:
            s += rev_m[m]
            cum_all[m] = s

# --- ставка НДС по консолидированному обороту на каждый месяц ---

        any_osno = any(cfg['orig_mode'] == 'ОСНО' for org, cfg in org_cfg.items() if cfg['consolidation'])

        nds_by_month = {}
        prev_gross = 0
        for m in months:
            curr_gross = cum_all[m]
            # если есть ОСНО в консолидации — всегда 20%, иначе считаем по шкале для "Доходы"
            mode_for_nds = 'ОСНО' if any_osno else 'Доходы'
            rateM = nds_rate(prev_gross, curr_gross, mode_for_nds, 0)
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

        p_rev, p_ebit, p_ebit_tax, p_net, last_mode = {}, {}, {}, {}, {}
        out = []
        usn_revoked_month = {}
        for g in records:
            cfg = org_cfg.get(g['org'], dict(orig_mode='ОСНО', consolidation=False,
                                            nds_rate=0, usn_rate=0, type='ООО'))
            key = 'consolidated' if cfg['consolidation'] else g['org']
            gross = cum_all[g['month']] if cfg['consolidation'] else cum_org[g['org']][g['month']]
            # --- логика перехода на ОСНО ---
            if (cfg['orig_mode'] in ('Доходы', 'Доходы-Расходы')
                and key not in usn_revoked_month
                and gross > LIMIT_GROSS_USN):
                usn_revoked_month[key] = g['month']
            if key in usn_revoked_month and g['month'] >= usn_revoked_month[key]:
                mode_eff = 'ОСНО'
            else:
                mode_eff = cfg['orig_mode']

            fot = fot_by_org.get(g['org'], 0)
            esn = esn_by_org.get(g['org'], 0)
            oklad_of = oklad_by_org.get(g['org'], 0)
            oth_cost = other.get(g['org'], 0)
            # дальше расчет показателей — НИКАКИХ пересчётов mode_eff и gross тут больше не нужно!



            
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
            mpNet   = mpGross / (1 + nds / 100)

            key = (g['org'], g['month'])
            fot = fot_by_org.get(g['org'], 0)
            esn = esn_by_org.get(g['org'], 0)
            oklad_of = oklad_by_org.get(g['org'], 0)


            oth_cost = other.get(g['org'], 0)

            if round(nds) in (5, 7):
                cost_base = full_cogs(g['cn'], nds)
            elif round(nds) == 20:
                cost_base = g['cn']
            else:
                cost_base = g['cr']

            cost_sales = cost_base
            cost_tax = g.get('ct', full_cogs(g['cn'], nds))
            cost_tax_wo = g.get('ct_wo', g['cn'])
            ebit_mgmt = revN - (cost_sales + mpNet + fot + esn + oth_cost)
            if mode_eff == 'Доходы-Расходы':
                ebit_tax = revN - (cost_tax + mpGross + oklad_of + esn + oth_cost)
                log_info(
                    f"[BASE] {g['org']} | m={g['month']:>02} | revN={revN:,.2f} - "
                    f"ct={cost_tax:,.2f} - mp={mpGross:,.2f} - of={oklad_of:,.2f} - "
                    f"esn={esn:,.2f} - oth={oth_cost:,.2f} = {ebit_tax:,.2f}"
                )

            else:
                ebit_tax = ebit_mgmt
            usn_base = ebit_tax

            # --- аккумулируем ---
            p_rev[g['org']] = p_rev.get(g['org'], 0) + g['rev']
            p_ebit[g['org']] = p_ebit.get(g['org'], 0) + ebit_mgmt
            p_ebit_tax[g['org']] = p_ebit_tax.get(g['org'], 0) + ebit_tax
            p_net[g['org']] = p_net.get(g['org'], 0) + revN

            out.append(dict(
                org=g['org'], m=g['month'], rev=g['rev'], cumG=gross,
                revN=revN, ndsSum=nds_sum, nds=nds,
                cr=g['cr'], cn=g['cn'], ct=cost_tax, ct_wo=cost_tax_wo,
                mpGross=mpGross, mpNet=mpNet,
                fot=fot, oklad_of=oklad_of, esn=esn, oth=oth_cost,
                ebit=ebit_mgmt,
                ebit_mgmt=ebit_mgmt,
                ebit_tax=ebit_tax,
                tax_base=usn_base,
                cumN=p_net[g['org']],
                cumE=p_ebit[g['org']],
                cumE_tax=p_ebit_tax[g['org']],
                mode=mode_eff, type=cfg['type'], prevM=last_mode.get(g['org']),
                usn=cfg['usn_rate'])
            )
            last_mode[g['org']] = mode_eff

        from collections import defaultdict

        # --- 1A. Консолидация "Доходы-Расходы" по месяцам ---
        consolidated_dr = [
            r for r in out
            if r['mode'] == 'Доходы-Расходы'
            and org_cfg.get(r['org'], {}).get('consolidation', False)
        ]

        cons_e_tax = _apply_consolidated_dr_tax(consolidated_dr)

        # 3. Для остальных организаций – по отдельности (за год 1–12)
        org_groups = defaultdict(list)
        for r in out:
            if r['mode'] == 'Доходы-Расходы' and not org_cfg.get(r['org'], {}).get('consolidation', False) and 1 <= r['m'] <= 12:
                org_groups[r['org']].append(r)

        for org, rows in org_groups.items():
            total_income = sum(r['revN'] for r in rows)

            # вместо суммы месячных tax_base берём cumE_tax последнего месяца
            last_row = max(rows, key=lambda x: x['m'])
            group_profit = max(last_row['cumE_tax'], 0)

            real_tax_sum = round(group_profit * rows[0]['usn'] / 100)
            min_tax_sum  = round(total_income * 0.01)
            use_min = real_tax_sum < min_tax_sum
            for r in rows:
                r['usn_forced_min'] = use_min
                r['tax'] = round(r['revN'] * 0.01) if use_min \
                           else round(max(r['tax_base'], 0) * r['usn'] / 100)

        # ---- 3A. Консолидация "Доходы" с учётом взносов по группе -----
        cons_income = defaultdict(list)
        for r in out:
            if (r['mode'] == 'Доходы' and
                    org_cfg.get(r['org'], {}).get('consolidation', False)):
                base = max(r['revN'], 0)
                raw_tax = round(base * r['usn'] / 100)
                r['raw_tax'] = raw_tax
                cons_income[r['m']].append(r)

        for m, rows in cons_income.items():
            total_raw = sum(r['raw_tax'] for r in rows)
            total_esn = sum(r['esn'] for r in rows)
            deduction_total = min(total_esn, total_raw * 0.5)
            for r in rows:
                share = r['raw_tax'] / total_raw if total_raw else 0
                r['deduction'] = round(deduction_total * share)

        ebit_m = acc(out, lambda x: x['m'], lambda x: x['ebit'])

        tax_base_cons_cum = {}
        run = 0
        for m in months:
            run += cons_e_tax.get(m, 0)
            tax_base_cons_cum[m] = run

        osno_cons_month = acc(
            (
                r
                for r in out
                if r['mode'] == 'ОСНО'
                and org_cfg.get(r['org'], {}).get('consolidation', False)
            ),
            lambda x: x['m'],
            lambda x: x['ebit_tax'],
        )
        osno_cons_cum = {}
        run = 0
        for m in months:
            run += osno_cons_month.get(m, 0)
            osno_cons_cum[m] = run
        # накопление прибыли по ОСНО: ключ 'consolidated' при консолидированном
        # учёте, иначе название организации
        rows_out, row_meta, cum_osno = [], [], {}
        for r in out:
            tax = base = 0
            rate = '0%'
            osno_cum = 0
            osno_cum_cons = ''
            if r['mode'] == 'Доходы':
                base = max(r['revN'], 0)
                raw_tax = r.get('raw_tax', round(base * r['usn'] / 100))
                if org_cfg.get(r['org'], {}).get('consolidation', False):
                    deduction = r.get('deduction', 0)
                else:
                    max_deduction = raw_tax * 0.5
                    deduction = min(r['esn'], max_deduction)
                tax = round(raw_tax - deduction)
                rate = f"{r['usn']}%"

                log_info(f"[TAX] {r['org']} | Доходы | base={base:,.2f} | raw={raw_tax} | esn={r['esn']} → tax={tax}")

            elif r['mode'] == 'Доходы-Расходы':
                if org_cfg.get(r['org'], {}).get('consolidation', False):
                    base = tax_base_cons_cum.get(r['m'], 0)
                    revenue = cum_all[r['m']]
                    rate_val = r['usn'] / 100
                    tax = round(calc_consolidated_min_tax(base, revenue, rate_val))
                    if tax == round(revenue * 0.01):
                        rate = '1%'
                        log_info(
                            f"[TAX] {r['org']} | Доходы-Расходы | применена "
                            f"минимальная ставка 1% от выручки ({tax:,.0f} ₽)"
                        )
                    else:
                        rate = f"{r['usn']}%"
                        log_info(
                            f"[TAX] {r['org']} | Доходы-Расходы | база={base:,.2f} "
                            f"| ставка={r['usn']}% → налог={tax}"
                        )
                else:
                    if r.get('usn_forced_min', False):
                        tax = round(r['revN'] * 0.01)
                        rate = '1%'
                    else:
                        base = max(r.get('tax_base', 0), 0)
                        tax = round(base * r['usn'] / 100)
                        rate = f"{(tax / base * 100):.2f}%" if base else '0%'
                    log_info(
                        f"[TAX] {r['org']} | Доходы-Расходы | base={base:,.2f} "
                        f"| tax={tax} | rate={rate}"
                    )
                    

            else:  # ОСНО
                if r['type'] == 'ИП':
                    # НДФЛ рассчитывается по накопленной прибыли.
                    # В режиме консолидации учитываем общий итог группы.
                    group_key = ('consolidated'
                                 if org_cfg.get(r['org'], {}).get('consolidation', False)
                                 else r['org'])
                    # Сбрасываем накопление только один раз при переходе группы
                    if r['prevM'] != 'ОСНО' and group_key not in cum_osno:

                        cum_osno[group_key] = 0

                    # --- ключ для накопления прибыли/убытка ---
                    # --- накопление полного EБIT (включая убытки) ---
                    prev = cum_osno.get(group_key, 0)
                    base = r['ebit_tax']
                    cum = prev + base

                    taxable_prev = max(prev, 0)
                    taxable_cum = max(cum, 0)
                    tax = max(0, round(
                        ndfl_prog(taxable_cum) - ndfl_prog(taxable_prev)
                    ))

                    cum_osno[group_key] = cum
                    osno_cum = cum_osno[group_key]
                    osno_cum_cons = (
                        osno_cons_cum.get(r['m'], 0)
                        if org_cfg.get(r['org'], {}).get('consolidation', False)
                        else ''
                    )

                    if osno_cum <= 0:
                        tax = 0
                        rate = '0%'
                        log_info(
                            f"[TAX] {r['org']} | ОСНО | group={group_key} | base={base:,.2f} → tax=0  (loss carry-forward)"
                        )
                    else:
                        rate = (
                            f"{(tax / max(base, 1) * 100):.2f}%" if base > 0 else '0%'
                        )
                        log_info(
                            f"[TAX] {r['org']} | ОСНО | group={group_key} | prev={prev:,.2f} | base={base:,.2f} → tax={tax}"
                        )
                else:
                    # Для юр. лиц ставка фиксированная, без накопления
                    base = max(r['ebit_tax'], 0)
                    tax = round(base * 0.25)
                    rate = '25%'
                    group_key = (
                        'consolidated'
                        if org_cfg.get(r['org'], {}).get('consolidation', False)
                        else r['org']
                    )
                    osno_cum = cum_osno.get(group_key, 0)
                    osno_cum_cons = (
                        osno_cons_cum.get(r['m'], 0)
                        if org_cfg.get(r['org'], {}).get('consolidation', False)
                        else ''
                    )
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
                # 11  Себестоимость Налог, ₽
                round(r['ct']),
                # 12  Себестоимость Налог без НДС, ₽
                round(r['ct_wo']),
                # 13  Расх. MP с НДС, ₽   (брутто)
                round(r['mpGross']),
                # 14  Расх. MP без НДС, ₽ (нетто)
                round(r['mpNet']),
                # 15  ФОТ, ₽
                round(r['fot']),
                # 16  Оклад_Оф, ₽
                round(r['oklad_of']),
                # 17  ЕСН, ₽
                round(r['esn']),
                # 18  Прочие, ₽
                round(r['oth']),
                # 19  EBITDA, ₽
                round(r['ebit_mgmt']),
                # 20  Расчет_базы_налога
                round(r['ebit_tax']),
                # 21  EBITDA нал. накоп., ₽
                round(r['cumE_tax']),
                # 22  EBITDA накоп., ₽
                round(r['cumE']),
                # 23  EBITDA сводно, ₽
                round(ebit_m[r['m']]),

                # 24  РасчетБазыНалогаНакопКонсол
                round(tax_base_cons_cum.get(r['m'], 0)),

                # 25  БазаНДФЛ ОСНО накоп., ₽
                round(osno_cum),

                # 26  БазаНДФЛ ОСНО накоп. сводно, ₽
                round(osno_cum_cons) if osno_cum_cons != '' else '',

                # 27  Режим
                r['mode'],
                # 28  Ставка УСН, %
                rate,
                # 29  Налог, ₽
                tax,
                # 30  Чистая прибыль, ₽
                round(r['ebit_mgmt'] - tax)
            ])

            row_meta.append(dict(
                org=r['org'],
                m=r['m'],
                mode=r['mode'],
                type=r['type'],
                consolidation=org_cfg.get(r['org'], {}).get('consolidation', False),
            ))

            if r['mode'] == 'ОСНО' and org_cfg.get(r['org'], {}).get('consolidation', False):
                log_info(
                    f"[OSNO CONS] {r['org']} | m={r['m']} | osno_cum={osno_cum:.2f} | osno_cum_cons={osno_cum_cons:.2f}"
                )
            else:
                log_info(
                    f"[OSNO INDV] {r['org']} | m={r['m']} | osno_cum={osno_cum:.2f} | osno_cum_cons=–"
                )

        consolidate_osno_tax(rows_out, row_meta)

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
                '[Red]_-* -#,##0 ₽_-;'     # отрицательные → красный минус
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

        for col in ruble_cols:
            idx = headers.index(col) + 1
            letter = col_name(idx)
            sh.range(total_row, idx).formula = \
                f"=SUBTOTAL(109,{letter}$2:{letter}${last_row})"
            sh.range(total_row, idx).number_format = fmt


        # ------ ярлык и позиция листа ----------------------------------
        sh.api.Tab.ColorIndex = 10
        if sh.index != 3:
            sh.api.Move(Before=ss.sheets[8].api)

        log_info(f'✔️  Готово! Записано строк: {len(rows_out)}')

    finally:
        if wb:
            wb.save()
            if app:
                wb.close()
                app.quit()

# ---------- 5. Точка входа -----------------------------------------------

if __name__ == '__main__':
    fill_planned_indicators()

