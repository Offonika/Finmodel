# file: calc_payroll_contributions.py
# -------------------------------------------------------------
# Расчёт страховых взносов, резерва отпусков и итоговой зарплаты
# (c) Финмодель • 2025-06-24
# -------------------------------------------------------------
# • Лист с сотрудниками  :  "ШтатноеРасписание"  (умная таблица StaffTbl)
# • Справочник льгот      :  "Справочник_льгот" (BenefitTbl)
# • Настройки организаций :  "НастройкиОрганизаций"
# • Параметры года        :  "Настройки"        (ParamTbl)
# -------------------------------------------------------------
# Работает и из Excel-макроса (RunPython) и из терминала.

# ---------- Пути --------------------------------------------------
import os, sys, logging, xlwings as xw, pandas as pd
from datetime import datetime
import pandas as pd
import datetime as dt
import argparse
from decimal import Decimal
import numpy as np

# -------- Показатели, которые будем выводить в сводный лист --------
SCENARIOS = {
    'as_is'        : 'Как есть',
    'all_white'    : 'Все белые',
    'optimal_white': 'Оптимум'
}
SUMMARY_FIELDS = ['Организация', 'Сценарий', 'ФОТ_белый', 'ФОТ_серый',
                  'Резерв', 'Взносы', '%_взносов',
                  'УСН_до', 'УСН_после', 'Нагрузка']

# -------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)           # ← папка Finmodel
EXCEL_PATH  = os.path.join(PROJECT_DIR, 'excel', 'Finmodel.xlsm')
LOG_DIR     = os.path.join(PROJECT_DIR, 'log')

os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(
    LOG_DIR, f'calc_payroll_{datetime.now():%Y%m%d_%H%M%S}.log'
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

parser = argparse.ArgumentParser(add_help=False)
parser.add_argument('--debug', action='store_true',
                    help='вывести расширенный DEBUG-лог во время расчёта')
ARGS, _ = parser.parse_known_args()  
log = logging.getLogger(__name__)
if ARGS.debug:
    log.setLevel(logging.DEBUG)
    for h in log.handlers:                   # чтобы DEBUG шёл и в файл, и в консоль
        h.setLevel(logging.DEBUG)
log.info('=== Запуск расчёта зарплаты ===')

# ---------- 1. Универсальное открытие --------------------------------------
def get_workbook():
    try:
        wb, app = xw.Book.caller(), None
        log.info('Запуск из Excel-макроса')
    except Exception:
        app = xw.App(visible=False)
        wb  = app.books.open(EXCEL_PATH)
        log.info('Запуск из терминала, открыт файл: %s', EXCEL_PATH)
    return wb, app

# ---------- 2. Служебные функции -------------------------------------------
def build_idx(header_row):
    """Словарь idx['Колонка'] = №_столбца (0-based)."""
    return {c: i for i, c in enumerate(header_row)}

# ---------- новая утилита ----------------------------------
def load_avg_revenue_netto(wb):
    """
    Из листа 'РасчетПлановыхПоказателей' берём колонку
    'Выручка без НДС, ₽' и считаем СРЕДНЕЕ за все строки
    (как правило — 12 месяцев) для каждой организации.
    Возвращает dict:  { 'Организация': avg_revenue_netto, ... }
    """
    sht = wb.sheets['РасчетПлановыхПоказателей']
    df  = (sht.range('A1')
              .options(pd.DataFrame, header=1,
                       index=False, expand='table')
              .value)

    if 'Организация' not in df.columns \
       or 'Выручка без НДС, ₽' not in df.columns:
        raise KeyError('В РасчетПлановыхПоказателей нет нужных колонок')

    df['Выручка без НДС, ₽'] = df['Выручка без НДС, ₽'].apply(to_float)
    df = df.dropna(subset=['Выручка без НДС, ₽'])      # убираем NaN/пустые
    # среднее без нулевых строк.  Если хотите учитывать нули — уберите .dropna()
    avg_rev = (df.groupby('Организация')['Выручка без НДС, ₽']
                 .mean()
                 .to_dict())

    log.debug('Средняя выручка (нетто) по организациям: %s',
              {k: round(v) for k, v in avg_rev.items()})
    return avg_rev


def to_float(val):
    """Преобразует число/проценты в float; всё остальное → None."""
    if pd.isna(val):
        return None

    # 1) Если это объект даты/времени — игнорируем
    if isinstance(val, (pd.Timestamp, dt.datetime, dt.date)):
        return None

    # 2) Строковые преобразования
    s = str(val).strip().replace(' ', '').replace('₽', '').replace(',', '.')
    if s in ('', '—', '-', 'Значение'):
        return None
    if s.endswith('%'):
        try:
            return float(s.rstrip('%')) / 100
        except ValueError:
            return None
    try:
        return float(s)
    except ValueError:
        return None           # всё, что не число, превращаем в None



# ---------- 3. Загрузка справочников ---------------------------------------
def load_parameters(wb):
    sht = wb.sheets['Настройки']

    # 1. Находим последнюю заполненную строку в колонке A
    last_row = sht.api.Cells(sht.api.Rows.Count, 1).End(-4162).Row   # xlUp = -4162

    # 2. Читаем диапазон A1:B<last_row> целиком
    df = sht.range((1, 1), (last_row, 2)).options(
            pd.DataFrame, header=0, index=False).value
    df.columns = ['Параметр', 'Значение']

    # 3. Чистим невидимые пробелы и отбрасываем пустые/заголовочные строки
    df['Параметр'] = (df['Параметр'].astype(str)
                      .str.replace(r'[\u00A0\u2007\u202F]', '', regex=True)
                      .str.strip())
    df = df[(df['Параметр'] != '') & (df['Параметр'] != 'Параметр')]

    # 4. Формируем словарь только из числовых значений
    params = {row['Параметр']: to_float(row['Значение'])
              for _, row in df.iterrows()
              if to_float(row['Значение']) is not None}
    for k, v in list(params.items()):
        if isinstance(v, Decimal): 
            params[k] = float(v)

    # DEBUG-контроль
    log.debug('Ключи params: %s', list(params.keys()))
    missing = [k for k in ('МРОТ', 'База_ПФР_предельная') if k not in params]
    if missing:
        raise KeyError(f'Отсутствуют параметры: {missing}')

    return params


def load_benefits(wb):
    sht = wb.sheets['Справочник_льгот']
    df  = sht.range('A1').options(pd.DataFrame, header=1,
                                  index=False, expand='table').value
    # очистка %-формата
    for col in ['Ставка до порога','Ставка сверх']:
        df[col] = df[col].apply(to_float)
    benefit = df.set_index('Категория_Льготы').to_dict('index')
    return benefit


# ----- вставьте вместо старой load_org_meta -----
def load_org_meta(wb):
    sht = wb.sheets['НастройкиОрганизаций']
    df  = (sht.range('A1')
              .options(pd.DataFrame, header=1, index=False, expand='table')
              .value)

    # чистим заголовки
    df.columns = (df.columns.astype(str)
                  .str.replace(r'[\u00A0\u2007\u202F]', '', regex=True)
                  .str.strip())

    # alias для старого имени колонки
    if 'РежимНалого' in df.columns and 'РежимНалогооблNew' not in df.columns:
        df.rename(columns={'РежимНалого': 'РежимНалогооблNew'}, inplace=True)

    cols = ['Организация', 'Категория_Льготы', 'Тариф_НСиПЗ',
            'РежимНалогооблNew', 'СтавкаНалогаУСН']

    meta = df.reindex(columns=cols).set_index('Организация')

    meta['Тариф_НСиПЗ']     = meta['Тариф_НСиПЗ'].apply(
                                lambda x: standardize_rate(to_float(x)))
    meta['СтавкаНалогаУСН'] = meta['СтавкаНалогаУСН'].apply(
                                lambda x: to_float(x) or 0.06)
    return meta
# -----------------------------------------------




def apply_scenario(df_org: pd.DataFrame,
                   mode: str,
                   *,
                   params: dict,
                   benefits: dict,
                   avg_rev_net: dict) -> pd.DataFrame:
    """
    df_org — DataFrame одной организации!
    mode   — 'as_is' | 'all_white' | 'optimal_white'
    """
    df = df_org.copy()

    # ── Простые режимы ─────────────────────────────────────────────
    if mode == 'as_is':
        return df

    if mode == 'all_white':
        df['Оклад_Оф'] = df['Оклад_Оф'].fillna(0) + df['Оклад_Серый'].fillna(0)
        df['Оклад_Серый'] = 0
        return df

    if mode != 'optimal_white':
        raise ValueError(f'Unknown mode {mode}')

    # ── Оптимизация только для УСН-«Доходы» ────────────────────────
    org    = df['Организация'].iloc[0]
    regime = str(df['РежимНалогооблNew'].iloc[0]).lower()

    if 'доходы' not in regime or 'расход' in regime:
        # не «Доходы» → вычета по взносам нет, оставляем как есть
        return df

    tax_rate = to_float(df['СтавкаНалогаУСН'].iloc[0]) or 0.06
    target_contrib = 0.5 * tax_rate * avg_rev_net.get(org, 0)
    if target_contrib == 0:
        return df

    mrot = params['МРОТ']

    # ── Подготовка массивов ────────────────────────────────────────
    W0 = df['Оклад_Оф'].fillna(0).astype(float).values
    G0 = df['Оклад_Серый'].fillna(0).astype(float).values

    # оценка взносов для одной строки
    def est_contrib(idx: int, white_val: float) -> float:
        row = df.iloc[idx]
        cat = row['Категория_Льготы'] or 'Без льготы'
        ben = benefits.get(cat, benefits['Без льготы'])
        lim = params['МРОТ'] * to_float(ben['Порог × МРОТ'])
        low = min(white_val, lim)
        high = max(0, white_val - lim)
        esv = low * to_float(ben['Ставка до порога']) + \
              high * to_float(ben['Ставка сверх'])
        nsipz = white_val * (row['Тариф_НСиПЗ'] or 0)
        return esv + nsipz

    # суммарные взносы при коэффициенте k
    def total_contrib(k: float) -> float:
        total = 0.0
        for i in range(len(W0)):
            white = max(mrot, W0[i] + k * G0[i])
            total += est_contrib(i, white)
        return total

    # ── Бинарный поиск k ∈ [0 ; 1] ─────────────────────────────────
    lo, hi = 0.0, 1.0
    for _ in range(25):              # точность < 1 руб.
        mid = (lo + hi) / 2
        if total_contrib(mid) > target_contrib:
            hi = mid
        else:
            lo = mid
    k = hi

    # если при k=1 взносы всё ещё ≤ target_contrib → всё делаем белым
    if total_contrib(1.0) <= target_contrib:
        k = 1.0

    # ── Применяем коэффициент ──────────────────────────────────────
    df['Оклад_Оф']    = np.maximum(W0 + k * G0, mrot).round()
    df['Оклад_Серый'] = (G0 * (1 - k)).round()

    return df


def standardize_rate(raw):
    if raw is None or pd.isna(raw):
        return 0.0
    r = float(raw)
    # ВСЁ, что > 1 %  ИЛИ  >= 1  трактуем как процент
    if r >= 1 or 0.01 < r <= 1:
        return r / 100
    return r          # уже доля ( <= 1 % )



# ---------- 4. Основной расчёт ---------------------------------------------
def calc_row(row, params, benedict):
    sal_off = float(row['Оклад_Оф']) if pd.notna(row['Оклад_Оф']) else 0
    sal_gray = float(row['Оклад_Серый']) if pd.notna(row['Оклад_Серый']) else 0


    # 4.1 Категория льготы
    cat = row['Категория_Льготы'] or 'Без льготы'
    ben = benedict.get(cat, benedict['Без льготы'])
    mrot_lim  = params['МРОТ'] * to_float(ben['Порог × МРОТ'])
    rate_low  = to_float(ben['Ставка до порога'])
    rate_high = to_float(ben['Ставка сверх'])
    part_low  = min(sal_off, mrot_lim)
    part_high = max(0, sal_off - mrot_lim)
    esv = part_low*rate_low + part_high*rate_high

    # 4.2 НСиПЗ
    risk_rate = row['Тариф_НСиПЗ'] or 0
    nsipz = sal_off * risk_rate
    percent_contrib = (esv + nsipz) / sal_off if sal_off else 0
    # 4.3 Резерв
    reserve_rate = float(params['Резерв_Отпусков_%'])
    reserve_pay  = sal_off * reserve_rate

    eff_rate = (esv + nsipz) / sal_off if sal_off else 0
    reserve_esv = reserve_pay * eff_rate
    # 4.4 Итоги
    total_vznosy   = esv + nsipz + reserve_esv
    total_payroll  = sal_off + sal_gray + reserve_pay + total_vznosy
    share_official = sal_off / (sal_off + sal_gray) if (sal_off+sal_gray) else 0

    # ------- DEBUG: детальная расшифровка -------
    log.debug(
        "ROW %s | org=%s | cat=%s | sal_off=%.2f | lim=%.2f | "
        "pLow=%.2f@%.3f  pHigh=%.2f@%.3f | "
        "ESV=%.2f  NSiPZ=%.2f  ResPay=%.2f  ResESV=%.2f  ==> Tot=%.2f",
        row.name,                     # индекс строки DataFrame
        row['Организация'],
        cat,
        sal_off,
        mrot_lim,
        part_low,  rate_low,
        part_high, rate_high,
        esv,
        nsipz,
        reserve_pay,
        reserve_esv,
        total_vznosy
    )
    # --------------------------------------------

    return pd.Series({
        'Доля_Оф'       : round(share_official, 3),
        '%_Взносов'   : round(percent_contrib, 4),
        'Итого_взносы'  : round(total_vznosy, 2),
        'Итого_зарплата': round(total_payroll, 2)
        
        
        #

    })

# ---------- 5. main() -------------------------------------------------------
def main():
    wb, app = get_workbook()
    try:
        avg_rev_net = load_avg_revenue_netto(wb)
        params   = load_parameters(wb)
        benefits = load_benefits(wb)
        org_meta = load_org_meta(wb)

        sht_staff = wb.sheets['ШтатноеРасписание']
        df = sht_staff.range('A1').options(pd.DataFrame, header=1,
                                           index=False, expand='table').value
        numeric_cols = ['Оклад_Оф', 'Оклад_Серый', 'Тариф_НСиПЗ']
        for c in numeric_cols:
            if c in df.columns:
                # to_float уже умеет убирать % и ₽, расширяем → float(...)
                df[c] = df[c].apply(lambda x: float(x) if isinstance(x, Decimal) else to_float(x))
        # mode = 'as_is' | 'all_white' | 'optimal_white'
    
        # (а) первый вызов  ───────────────────────────────────────────────
        df = apply_scenario(
                df,                         # тот же DataFrame
                mode='as_is',
                params=params,
                benefits=benefits,
                avg_rev_net=avg_rev_net
        )

       

        if df.empty:
            log.warning('Таблица ШтатноеРасписание пуста, расчёт прерван')
            return

        # 5.1 Подставляем недостающие льготы/тарифы из мета-таблицы
        df = df.merge(org_meta, how='left', left_on='Организация',
                      right_index=True, suffixes=('','_org'))
        df['Категория_Льготы'] = df['Категория_Льготы']\
            .fillna(df['Категория_Льготы_org'])
        df['Тариф_НСиПЗ'] = df['Тариф_НСиПЗ']\
            .fillna(df['Тариф_НСиПЗ_org'])
        df.drop(columns=['Категория_Льготы_org','Тариф_НСиПЗ_org'],
                inplace=True)
        df['Тариф_НСиПЗ'] = df['Тариф_НСиПЗ']\
            .apply(lambda x: standardize_rate(to_float(x)))
        # 5.2 Построчный расчёт
        calc = df.apply(calc_row, axis=1, args=(params, benefits))
        df[['Доля_Оф','%_Взносов','Итого_взносы','Итого_зарплата']] = calc
        


        # ── 5.4 Формируем сводку по сценариям ──────────────────────────
        summary_rows: list[list] = []
        scen_rows:    list[pd.DataFrame] = []

# ---------- utils / numeric ------------------------------------------------
        from functools import lru_cache
         
        @lru_cache(maxsize=None)                # чтобы предупреждение писалось один раз
        def _warn_once(cols_signature: tuple[str, ...]) -> None:
            """
            Логируем повторяющиеся колонки ровно один раз для каждого уникального
            набора имён.  Используем lru_cache → работает как «один-разовый» флаг.
            """
            log.debug(
                "[WARN] duplicate columns: %s — берём только первую",
                list(cols_signature)
            )

        def to_numeric(col_or_df: 'pd.Series | pd.DataFrame') -> pd.Series:
            """
            Универсальный перевод числовой колонки (или DataFrame-дубликата)
            в float-Series.

            * чистит пробелы, «₽», NBSP, запятые → точки;
            * ошибки → 0.0;
            * если подаётся DataFrame с дублирующимся именем колонки,
            оставляем **первый** столбец и выводим предупреждение (один раз).
            """
            # ―― 1. выбираем столбец ―――――――――――――――――――――――――――――――――――――――――――――――
            if isinstance(col_or_df, pd.DataFrame):
                if col_or_df.shape[1] > 1:
                    _warn_once(tuple(col_or_df.columns))
                col = col_or_df.iloc[:, 0]
            else:
                col = col_or_df

            # ―― 2. чистим и конвертируем ――――――――――――――――――――――――――――――――――――――――
            cleaned = (
                col.astype(str)
                .str.replace(r'[₽\s\u00A0]', '', regex=True)   # «₽» и все пробелы
                .str.replace(',', '.', regex=False)            # запятая → точка
            )
            return pd.to_numeric(cleaned, errors='coerce').fillna(0.0)
        # ---------- utils / numeric (end) ------------------------------------------
        for org, df_org in df.groupby('Организация'):
            df_org_base = df_org.copy()

            # 1. Обогащаем данными из справочника (до расчёта сценария!)
            d = (df_org_base.copy()
                .merge(org_meta, how='left', left_on='Организация', right_index=True, suffixes=('', '_org'))
                .assign(
                    Категория_Льготы=lambda x: x['Категория_Льготы'].fillna(x['Категория_Льготы_org']),
                    Тариф_НСиПЗ=lambda x: x['Тариф_НСиПЗ'].fillna(x['Тариф_НСиПЗ_org']).apply(lambda r: standardize_rate(to_float(r))),
                    РежимНалогооблNew=lambda x: x['РежимНалогооблNew'].fillna(x['РежимНалогооблNew_org']),
                    СтавкаНалогаУСН=lambda x: x['СтавкаНалогаУСН'].fillna(x['СтавкаНалогаУСН_org'])
                )
                .drop(columns=['Категория_Льготы_org', 'Тариф_НСиПЗ_org', 'РежимНалогооблNew_org', 'СтавкаНалогаУСН_org'])
            )

            for mode, title in SCENARIOS.items():
                # 2. Применяем сценарий к уже обогащённому df
                d_mode = apply_scenario(
                    d.copy(),
                    mode,
                    params=params,
                    benefits=benefits,
                    avg_rev_net=avg_rev_net
                )

                # 3. Пересчёт построчно
                calc_parts = d_mode.apply(calc_row, axis=1, args=(params, benefits))
                d_mode = d_mode.assign(**calc_parts)

                # 4. Получаем актуальную ставку налога из d_mode
                tax_rate = to_float(d_mode['СтавкаНалогаУСН'].iloc[0]) or 0.06

                white   = to_numeric(d_mode['Оклад_Оф']).sum()
                gray    = to_numeric(d_mode['Оклад_Серый']).sum()
                reserv  = (to_numeric(d_mode['Оклад_Оф']) * params['Резерв_Отпусков_%']).sum()
                contrib = to_numeric(d_mode['Итого_взносы']).sum()

                full_payroll = white + gray
                tax_base     = avg_rev_net.get(org, 0)

                tax0 = tax_rate * tax_base
                pct  = contrib / full_payroll if full_payroll else 0

                if 'доходы' in str(d_mode['РежимНалогооблNew'].iloc[0]).lower() and 'расход' not in str(d_mode['РежимНалогооблNew'].iloc[0]).lower():
                    tax1 = max(tax0 - contrib, tax0 * 0.5)
                else:
                    tax1 = tax0

                total = white + gray + reserv + contrib

                summary_rows.append([
                    org, title,
                    round(white), round(gray), round(reserv),
                    round(contrib), round(pct * 100, 2),
                    round(tax0),   round(tax1), round(total)
                ])

                scen_rows.append(
                    d_mode.assign(Сценарий=title)[
                        ['Организация', 'ФИО',
                        'Оклад_Оф', 'Оклад_Серый', 'Доля_Оф',
                        'Категория_Льготы', 'Тариф_НСиПЗ', '%_Взносов',
                        'Итого_взносы', 'Итого_зарплата', 'Сценарий']
                    ]
                )


        # ── 6. Запись сводных листов ───────────────────────────────────
 

        # зелёный ярлык и позиция
        try:
            sht_sum.api.Tab.ColorIndex = 35                  # зелёная вкладка
            if sht_sum.index != 4:                           # хотим четвёртым
                sht_sum.api.Move(Before=wb.sheets[3].api)
        except Exception:
            pass

        # 6-B. «РасчетЗарплаты» — деталка по сотрудникам (если нужна)
        sheet_name_det = 'РасчетЗарплаты'
        sht_det = (wb.sheets[sheet_name_det]
                   if sheet_name_det in [s.name for s in wb.sheets]
                   else wb.sheets.add(sheet_name_det))
        sht_det.clear()

        if scen_rows:
            scen_df = pd.concat(scen_rows, ignore_index=True)
            sht_det.range(1, 1).value = scen_df

            # --- СОЗДАЁМ УМНУЮ ТАБЛИЦУ (ListObject) ---
            table_range = sht_det.range('A1').expand()
            # удаляем старые таблицы, если были (иначе ошибка)
            for tbl in sht_det.api.ListObjects:
                tbl.Delete()
            sht_det.api.ListObjects.Add(
                SourceType=1,               # xlSrcRange
                Source=table_range.api,
                XlListObjectHasHeaders=1
            ).Name = 'ScenPayrollTbl'       # имя можно своё

            # Применяем стиль "TableStyleMedium7" (зелёный средний)
            sht_det.api.ListObjects('ScenPayrollTbl').TableStyle = "TableStyleMedium7"

            sht_det.range('A1').expand().columns.autofit()
            sht_det.range('1:1').font.bold = True





    # ---------- обработка ошибок и закрытие Excel -------------------------
    except Exception as e:
        log.exception('Ошибка расчёта: %s', e)
        raise                              # пробросим дальше – пусть видно в консоли
    finally:
        if app is not None:                # закрываем только если файл открывали мы
            wb.save()
            app.quit()
            log.info('Excel-файл сохранён и закрыт')
        log.info('=== Конец расчёта ===')


if __name__ == '__main__':
    main()