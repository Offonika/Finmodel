from pathlib import Path
import xlwings as xw
import requests
import datetime
import time
from collections import Counter
import pandas as pd

SHEET_SETTINGS  = "Настройки"
SHEET_ORGS      = "НастройкиОрганизаций"
SHEET_FACTS     = "ФинотчетыWB"
SHEET_LOG       = "WB_Log"

WB_API_URL_STAT = "https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod"
BATCH_SIZE_WB   = 100000

EXCEL_FILE_PATH = Path(__file__).resolve().parents[1] / "Finmodel.xlsm"

def get_idx(header):
    return {str(h).strip(): i for i, h in enumerate(header)}

def log_step(ws_log, msg):
    print("[INFO]", msg)
    last_row = ws_log.range('A' + str(ws_log.cells.rows.count)).end('up').row
    ws_log.range(f'A{last_row + 1}').value = [datetime.datetime.now(), "INFO", msg]

def log_error(ws_log, msg):
    print("[ERROR]", msg)
    last_row = ws_log.range('A' + str(ws_log.cells.rows.count)).end('up').row
    ws_log.range(f'A{last_row + 1}').value = [datetime.datetime.now(), "ERROR", msg]

def get_or_create_sheet(wb, sheet_name, ws_log=None):
    if sheet_name in [s.name for s in wb.sheets]:
        ws = wb.sheets[sheet_name]
    else:
        ws = wb.sheets.add(sheet_name)
        if ws_log is not None:
            log_step(ws_log, f"Создан новый лист: {sheet_name}")
    return ws

def fetch_wb_report_stat(token, date_from, date_to, rrdid, ws_log=None, org=""):
    rr_part = f"&rrdid={int(float(rrdid))}" if rrdid is not None else ""
    url = f"{WB_API_URL_STAT}?dateFrom={date_from}&dateTo={date_to}&limit={BATCH_SIZE_WB}{rr_part}"

    if ws_log is not None:
        log_step(ws_log, f"→ [WB API] {org} | Запрос: {url[:120]}...")
    print(f"→ [WB API] {org} | Запрос: {url}")

    headers = {"Authorization": token}
    try:
        resp = requests.get(url, headers=headers, timeout=90)
        print(f"← [WB API] {org} | status={resp.status_code} | bytes={len(resp.content)}")
        if resp.status_code == 429:
            log_error(ws_log, f"[WB API] {org} | Статус 429 Too Many Requests. Ожидание 60 секунд...")
            time.sleep(60)
            return [], rrdid
        if resp.status_code != 200:
            log_error(ws_log, f"[WB API] {org} | Код ответа: {resp.status_code}. Пропуск итерации.")
            return [], rrdid
        data = resp.json()
        if not isinstance(data, list):
            log_error(ws_log, f"[WB API] {org} | Ответ не является списком!")
            return [], rrdid
        last_rrd = data[-1]['rrd_id'] if data else rrdid
        return data, last_rrd
    except requests.exceptions.Timeout:
        log_error(ws_log, f"[WB API] {org} | Таймаут запроса! Пропуск итерации.")
        return [], rrdid
    except Exception as e:
        log_error(ws_log, f"[WB API] {org} | Ошибка: {type(e).__name__}: {e}")
        return [], rrdid

def aggregate_wb_rows(rows, org, doc_types_counter=None):
    agg = {}
    for r in rows:
        doc_type = (r.get('doc_type_name') or "").strip().lower()
        if doc_types_counter is not None:
            doc_types_counter[doc_type] += 1
        key = (
            org,
            r.get('realizationreport_id'),
            r.get('nm_id'),
            r.get('sa_name')
        )
        if key not in agg:
            agg[key] = {
                'Организация': org,
                'Дата': (r.get('create_dt') or "")[:10],
                'Предмет': r.get('subject_name') or "",
                'Артикул_продавца': r.get('sa_name') or "",
                'Артикул_WB': r.get('nm_id') or "",
                'Название': r.get('brand_name') or "",
                'Номер_отчёта': r.get('realizationreport_id') or "",
                'Продано_шт': 0, 'Возврат_шт': 0,
                'Продано_руб': 0, 'Возвраты_руб': 0,
                'Выручка': 0, 'К_перечислению_за_товар': 0, 'Комиссия': 0,
                'Стоимость_логистики': 0, 'Стоимость_хранения': 0,
                'Стоимость_платной_приемки': 0, 'Общая_сумма_штрафов': 0,
                'Прочие_удержания_выплаты': 0, 'Доплаты': 0,
                'Итого_к_оплате': 0, 'Итого_продано': 0
            }
        a = agg[key]
        qty   = float(r.get('quantity') or 0)
        retail = float(r.get('retail_amount') or 0)
        ppvz   = float(r.get('ppvz_for_pay') or 0)

        if doc_type.startswith("прода"):
            a['Продано_шт'] += qty
            a['Продано_руб'] += retail
            a['К_перечислению_за_товар'] += ppvz
        elif doc_type.startswith("возвра"):
            a['Возврат_шт'] += qty
            a['Возвраты_руб'] += retail
            a['К_перечислению_за_товар'] -= ppvz

        a['Стоимость_логистики']       += float(r.get('delivery_rub') or 0)
        a['Стоимость_хранения']        += float(r.get('storage_fee') or 0)
        a['Стоимость_платной_приемки'] += float(r.get('acceptance') or 0)
        a['Общая_сумма_штрафов']       += float(r.get('penalty') or 0)
        a['Прочие_удержания_выплаты']  += float(r.get('deduction') or 0)
        a['Доплаты']                   += float(r.get('additional_payment') or 0)

    for a in agg.values():
        a['Итого_продано']  = a['Продано_шт'] - a['Возврат_шт']
        a['Выручка']        = a['Продано_руб'] - a['Возвраты_руб']
        a['Комиссия']       = a['Выручка'] - a['К_перечислению_за_товар']
        a['Итого_к_оплате'] = (
            a['К_перечислению_за_товар']
            - a['Стоимость_логистики'] - a['Стоимость_хранения']
            - a['Стоимость_платной_приемки'] - a['Общая_сумма_штрафов']
            - a['Прочие_удержания_выплаты'] + a['Доплаты']
        )
    return list(agg.values())

def parse_any_date(dt_str):
    if isinstance(dt_str, (datetime.datetime, datetime.date)):
        return dt_str
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.datetime.strptime(dt_str, fmt)
        except Exception:
            continue
    return None

def drop_existing_table(ws, table_name="WbFactsTable"):
    for tbl in ws.tables:
        if tbl.name == table_name:
            tbl.api.Delete()
            return

def _norm(val: str) -> str:
    """
    Нормализует числовые/строковые коды:
    • убирает пробелы и символы перевода строки;
    • превращает 123456789.0 → 123456789;
    • всегда возвращает str.
    """
    return str(val).strip().split('.')[0]

def split_periods_by_week(date_from, date_to):
    start = pd.to_datetime(date_from)
    end   = pd.to_datetime(date_to)
    periods = []
    cur = start
    while cur <= end:
        week_start = cur
        week_end   = min(end, cur + pd.Timedelta(days=6))
        periods.append((week_start.strftime("%Y-%m-%d"), week_end.strftime("%Y-%m-%d")))
        cur = week_end + pd.Timedelta(days=1)
    return periods

def import_wb_detailed_reports(wb=None):
    if wb is None:
        try:
            wb = xw.Book.caller()
        except Exception:
            wb = xw.Book(str(EXCEL_FILE_PATH))

    ws_log   = get_or_create_sheet(wb, SHEET_LOG)
    ws_set   = get_or_create_sheet(wb, SHEET_SETTINGS,  ws_log)
    ws_org   = get_or_create_sheet(wb, SHEET_ORGS,      ws_log)
    ws_facts = get_or_create_sheet(wb, SHEET_FACTS,     ws_log)

    set_data = ws_set.range('A1').expand().value
    idx_set  = get_idx(set_data[0])
    params   = {str(r[idx_set['Параметр']]).strip(): str(r[idx_set['Значение']]).strip()
                for r in set_data[1:] if r and len(r) >= 2}

    date_from = params.get("ПериодНачало")
    date_to   = params.get("ПериодКонец")
    if not date_from or not date_to:
        log_error(ws_log, "Нет ПериодНачало или ПериодКонец в Настройках!")
        return

    date_from_iso = parse_any_date(date_from).strftime("%Y-%m-%d")
    date_to_iso   = parse_any_date(date_to).strftime("%Y-%m-%d")

    org_data  = ws_org.range('A1').expand().value
    idx_org   = get_idx(org_data[0])
    idx_rrdid = idx_org.get('rrd_id')

    old_facts = ws_facts.range('A1').expand().value
    existing_keys = set()
    if old_facts and len(old_facts) > 1:
        idx_f = get_idx(old_facts[0])
        for r in old_facts[1:]:
            if not r:
                continue
            key = (
                _norm(r[idx_f['Организация']]),
                _norm(r[idx_f['Номер_отчёта']]),
                _norm(r[idx_f['Артикул_WB']]),
                _norm(r[idx_f['Артикул_продавца']])
            )
            existing_keys.add(key)


            
    log_step(ws_log, f"Найдены ранее загруженные строки: {len(existing_keys)}")

    results = []
    doc_types_counter = Counter()

    for row_idx, row in enumerate(org_data[1:], start=2):
        org   = row[idx_org['Организация']]
        token = row[idx_org['Token_WB']]
        if not org or not token:
            continue

        log_step(ws_log, f"Обработка: {org}")
        periods = split_periods_by_week(date_from_iso, date_to_iso)
        log_step(ws_log, f"Загрузка разбита на {len(periods)} недель: {periods}")
        print(f"Загрузка разбита на {len(periods)} недель: {periods}")

        for period_start, period_end in periods:
            local_rrd = 0
            max_rrd = 0
            attempts = 0
            while True:
                attempts += 1
                if attempts > 5:
                    log_error(ws_log, f"Превышено число попыток (5) для {org}, {period_start}—{period_end}. Прерывание.")
                    break

                wb_rows, last_rrd = fetch_wb_report_stat(
                    token, period_start, period_end, local_rrd, ws_log, org
                )

                log_step(ws_log, f"[{org}] {period_start}—{period_end} | Получено строк: {len(wb_rows)} (попытка {attempts})")
                print(f"[{org}] {period_start}—{period_end} | Получено строк: {len(wb_rows)} (попытка {attempts})")

                if not wb_rows:
                    break

                records_batch = aggregate_wb_rows(wb_rows, org, doc_types_counter)
                records_new = []
                for rec in records_batch:
                    key = (
                        _norm(rec['Организация']),
                        _norm(rec['Номер_отчёта']),
                        _norm(rec['Артикул_WB']),
                        _norm(rec['Артикул_продавца'])
                    )


                    if key in existing_keys:
                        continue
                    existing_keys.add(key)
                    records_new.append(rec)

                if records_new:
                    results.extend(records_new)

                if last_rrd > max_rrd:
                    max_rrd = last_rrd
                if last_rrd == local_rrd:
                    break
                local_rrd = last_rrd
                time.sleep(1)

            # --- обновление rrd_id после недели ---
            if idx_rrdid is not None and max_rrd > int(float(row[idx_rrdid] or 0)):
                ws_org.range(row_idx, idx_rrdid + 1).value = max_rrd
                log_step(ws_log, f"Обновлён rrd_id для {org} (неделя): {max_rrd}")

    if not results:
        log_step(ws_log, "Нет новых строк — лист ФинотчетыWB оставлен без изменений")
        return

    hdr = [
        'Организация', 'Дата', 'Предмет', 'Артикул_продавца', 'Артикул_WB', 'Название', 'Номер_отчёта',
        'Продано_шт', 'Возврат_шт', 'Продано_руб', 'Возвраты_руб', 'Выручка', 'К_перечислению_за_товар', 'Комиссия',
        'Стоимость_логистики', 'Стоимость_хранения', 'Стоимость_платной_приемки', 'Общая_сумма_штрафов',
        'Прочие_удержания_выплаты', 'Доплаты', 'Итого_к_оплате', 'Итого_продано'
    ]

    def header_ok(row): return all(col in row for col in hdr)

    if old_facts and header_ok(old_facts[0]):
        start = len(old_facts) + 1
        for i, rec in enumerate(results, start):
            ws_facts.range(i, 1).value = [rec.get(c, "") for c in hdr]
    else:
        drop_existing_table(ws_facts, "WbFactsTable")
        ws_facts.clear_contents()
        ws_facts.range('A1').value = hdr
        for i, rec in enumerate(results, 2):
            ws_facts.range(i, 1).value = [rec.get(c, "") for c in hdr]

    log_step(ws_log, f"Финально записано {len(results)} новых строк в {SHEET_FACTS}")

    try:
        ws_facts.activate()
        last_row = ws_facts.range('A' + str(ws_facts.cells.rows.count)).end('up').row
        last_col = len(hdr)
        tbl_rng  = ws_facts.range((1, 1), (last_row, last_col))

        existing_tbl = None
        for tbl in ws_facts.tables:
            if tbl.name == "WbFactsTable":
                existing_tbl = tbl
                break

        if existing_tbl:
            existing_tbl.resize(tbl_rng)
        else:
            ws_facts.tables.add(tbl_rng,
                                name="WbFactsTable",
                                table_style_name="TableStyleMedium7",
                                has_headers=True)

        ws_facts.api.Tab.Color = 0xEAF884
        if ws_facts.index != 13:
            ws_facts.api.Move(Before=wb.sheets[12].api)

        log_step(ws_log, "✓ Лист оформлен/обновлён как умная таблица")

    except Exception as e:
        log_error(ws_log, f"⚠️ Ошибка оформления таблицы: {e}")

if __name__ == "__main__":
    import_wb_detailed_reports()
