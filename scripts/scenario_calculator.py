# scenario_calculator.py – «Что‑если»-модуль к fill_planned_indicators.py

import argparse
import os
import xlwings as xw
from pathlib import Path
from collections import defaultdict
from fill_planned_indicators import (
    open_wb,             # открыть/подсоединиться к Excel-книге
    parse_money, parse_month,
    nds_rate, ndfl_prog,
    build_idx, read_rows
)

def normalize(s):
    """Нормализация заголовков: убрать пробелы, привести к нижнему регистру, заменить _."""
    return str(s).strip().lower().replace(" ", "").replace("_", "")

# CLI
PAR = argparse.ArgumentParser(description="Scenario profit calculator")
PAR.add_argument("-f", "--file", default="Finmodel.xlsm", help="Excel workbook")
ARGS = PAR.parse_args()

IS_EXE = getattr(sys, "frozen", False)
BASE_DIR = Path(sys.executable if IS_EXE else __file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent

EXCEL_PATH = PROJECT_DIR / ARGS.file

if not EXCEL_PATH.exists():
    raise FileNotFoundError(f"Workbook not found: {EXCEL_PATH}")

# Листы
SHEET_WB   = "РасчётЭкономикиWB"
SHEET_OZON = "РасчетЭкономикиОзон"
SHEET_ORG  = "НастройкиОрганизаций"
SHEET_SAL  = "Зарплата"
SHEET_OTH  = "ПрочиеРасходы"
SHEET_RES  = "Сценарии"

LIMIT_GROSS_USN = 450_000_000

def load_inputs(wb):
    sheet_names = [s.name for s in wb.sheets]
    raw = []
    # Чтение WB и OZON
    if SHEET_WB in sheet_names:
        rows, idx = read_rows(wb.sheets[SHEET_WB])
        raw += [ [*idx.keys()] ] + rows
    if SHEET_OZON in sheet_names:
        rows, idx = read_rows(wb.sheets[SHEET_OZON])
        raw += [ [*idx.keys()] ] + rows

    # Настройки организаций (с заголовком!)
    cfg_rows = []
    if SHEET_ORG in sheet_names:
        rows, idx = read_rows(wb.sheets[SHEET_ORG])
        cfg_rows = [list(idx.keys())] + rows

    # Зарплата
    sal_rows = []
    if SHEET_SAL in sheet_names:
        rows, idx = read_rows(wb.sheets[SHEET_SAL])
        sal_rows = [list(idx.keys())] + rows

    # Прочие расходы
    oth_rows = []
    if SHEET_OTH in sheet_names:
        rows, idx = read_rows(wb.sheets[SHEET_OTH])
        oth_rows = [list(idx.keys())] + rows

    return raw, cfg_rows, sal_rows, oth_rows

def group_records(raw):
    header = [normalize(c) for c in raw[0]]
    idx = {h: i for i, h in enumerate(header)}
    data = raw[1:]
    groups = {}
    for row in data:
        if normalize('месяц') not in idx or normalize('организация') not in idx:
            continue
        month = parse_month(row[idx[normalize('месяц')]])
        if not isinstance(month, int) or not (1 <= month <= 12):
            continue
        org = row[idx[normalize('организация')]]
        rev = parse_money(row[idx.get(normalize('выручка, ₽'), idx.get(normalize('выручка'), 0))])
        mp  = parse_money(row[idx.get(normalize('расходы мп, ₽'), idx.get(normalize('расходы мп'), 0))])
        cr  = parse_money(row[idx.get(normalize('себестоимостьпродажруб'), '')])
        cn  = parse_money(row[idx.get(normalize('себестоимостьпродажбезндс'), '')])
        key = (org, month)
        if key not in groups:
            groups[key] = dict(org=org, month=month, rev=0, mp=0, cr=0, cn=0)
        g = groups[key]
        g['rev'] += rev; g['mp']  += mp
        g['cr']  += cr;  g['cn'] += cn
    return list(groups.values())

def make_cfg_dict(cfg_rows):
    if not cfg_rows:
        return {}
    header = [normalize(h) for h in cfg_rows[0]]
    idx = {h: i for i, h in enumerate(header)}
    cfg_dict = {}
    for row in cfg_rows[1:]:
        org = row[idx[normalize('организация')]]
        cfg_dict[org] = {
            'orig_mode': str(row[idx.get(normalize('режимналогооблnew'), '')]).strip() or 'ОСНО',
            'consolidation': str(row[idx.get(normalize('консолидация'), '')]).strip().lower() == 'да',
            'nds_rate': parse_money(str(row[idx.get(normalize('ставка ндс'), '')]).replace('%', '').replace(',', '.')),
            'usn_rate': parse_money(str(row[idx.get(normalize('ставканалогаусн'), '')]).replace('%', '').replace(',', '.')),
            'type': str(row[idx.get(normalize('тип_организации'), '')]).strip() or 'ООО'
        }
    return cfg_dict

def make_salary_dict(sal_rows):
    if not sal_rows:
        return {}
    header = [normalize(h) for h in sal_rows[0]]
    idx = {h: i for i, h in enumerate(header)}
    salary_dict = {}
    for row in sal_rows[1:]:
        org = row[idx[normalize('организация')]]
        salary_dict[org] = {
            'fot': parse_money(row[idx.get(normalize('фот'), '')]),
            'mode': str(row[idx.get(normalize('режим_зп'), '')]).strip()
        }
    return salary_dict

def make_other_dict(oth_rows):
    if not oth_rows:
        return {}
    header = [normalize(h) for h in oth_rows[0]]
    idx = {h: i for i, h in enumerate(header)}
    other_dict = {}
    for row in oth_rows[1:]:
        org = row[idx[normalize('организация')]]
        val = parse_money(row[idx.get(normalize('прочие'), 1)])  # если нет — 1-я колонка
        other_dict[org] = val
    return other_dict

def calc_scenario(records, cfg, salary, other,
                  consolidate_all=False,
                  forced_mode=None,
                  min_nds=None):
    months   = sorted({r['month'] for r in records})
    cum_all  = defaultdict(float)
    for m in months:
        cum_all[m] = sum(r['rev'] for r in records if r['month'] <= m)

    cum_org  = defaultdict(lambda: defaultdict(float))
    for org in {r['org'] for r in records}:
        run = 0
        for m in months:
            run += sum(r['rev'] for r in records if r['org']==org and r['month']==m)
            cum_org[org][m] = run

    res = defaultdict(lambda: dict(ebitda=0, tax=0, profit=0))
    p_rev, p_ebit = defaultdict(float), defaultdict(float)
    nds_by_m = {}
    prev_g = 0
    for m in months:
        nds_by_m[m] = nds_rate(prev_g, cum_all[m], 'Доходы', 0)
        prev_g = cum_all[m]

    for r in records:
        if r['org'] not in cfg:
            continue
        c = cfg[r['org']].copy()
        if forced_mode:
            c['orig_mode'] = forced_mode
        if consolidate_all:
            c['consolidation'] = True
        if min_nds is not None:
            c['nds_rate'] = max(c['nds_rate'], min_nds)

        gross = cum_all[r['month']] if c['consolidation'] else cum_org[r['org']][r['month']]
        mode  = c['orig_mode']
        if mode in ('Доходы','Доходы-Расходы') and gross > LIMIT_GROSS_USN:
            mode='ОСНО'
        if c['consolidation']:
            nds = nds_by_m[r['month']]
        else:
            prev = p_rev[r['org']]
            nds  = nds_rate(prev, prev+r['rev'], mode, c['nds_rate'])
        nds = max(nds, c['nds_rate'])

        revN = r['rev']/(1+nds/100) if nds else r['rev']
        mpN  = r['mp']/1.2 if r['mp'] else 0
        cost = r['cn'] if round(nds)==20 else r['cr']
        sal = salary.get(r['org'], {'fot':0, 'mode':''})
        esn = sal['fot']*0.3 if sal['mode']=='Официальная' else 0
        ebit = revN - (cost + mpN + sal['fot'] + esn + other.get(r['org'],0))
        p_rev[r['org']]  += r['rev']
        p_ebit[r['org']] += ebit

        # Налоги
        tax = 0
        if mode=='Доходы':
            tax = max(revN,0)*c['usn_rate']/100
        elif mode=='Доходы-Расходы':
            tax = max(ebit,0)*c['usn_rate']/100
        elif mode=='ОСНО':
            if c['type']=='ИП':
                tax = max(ebit,0)*0.13
            else:
                tax = max(ebit,0)*0.25

        res[r['org']]['ebitda'] += ebit
        res[r['org']]['tax']    += tax
        res[r['org']]['profit'] += ebit-tax

    total = sum(v['profit'] for v in res.values())
    return total, res

def main():
    wb, app = None, None
    try:
        wb, app = open_wb()
        raw, cfg_rows, sal_rows, oth_rows = load_inputs(wb)
        if not raw:
            print("Нет данных для анализа!")
            return
        records     = group_records(raw)
        cfg_dict    = make_cfg_dict(cfg_rows)
        salary_dict = make_salary_dict(sal_rows)
        other_dict  = make_other_dict(oth_rows)

        scenarios = [
            dict(name='Текущие настройки', consolidate_all=False, forced_mode=None),
            dict(name='Все ОСНО (без конс.)', consolidate_all=False, forced_mode='ОСНО'),
            dict(name='Доходы (конс.)',      consolidate_all=True,  forced_mode='Доходы'),
            dict(name='Доходы‑Расходы (конс.)', consolidate_all=True, forced_mode='Доходы-Расходы'),
        ]

        summary = []
        for sc in scenarios:
            total, by_org = calc_scenario(
                records, cfg_dict, salary_dict, other_dict,
                consolidate_all=sc.get('consolidate_all', False),
                forced_mode=sc.get('forced_mode'),
                min_nds=None
            )
            summary.append((sc['name'], total))
            print(f"{sc['name']:<30} → {total:>15,.0f} ₽")

        # Запишем результат в Excel
        if SHEET_RES in [s.name for s in wb.sheets]:
            wb.sheets[SHEET_RES].delete()
        sh = wb.sheets.add(SHEET_RES)
        sh.range(1, 1).value = ['Сценарий', 'Чистая прибыль, ₽']
        sh.range(2, 1).value = summary
        sh.range(2, 2).number_format = '#,##0 ₽'

        print(f"✔️ Итог записан на лист '{SHEET_RES}'")
    finally:
        if wb:
            wb.save()
            if app: app.quit()

if __name__ == '__main__':
    main()
