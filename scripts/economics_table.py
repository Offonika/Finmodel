# economics_table.py
# ------------------------------------------------------------------
# Создание таблицы «РасчетЭкономикиОзон» по данным листов
# «ПланПродажОзон», «РасчётСебестоимости» и «Настройки».
# Скрипт работает как из Excel-макроса (RunPython), так и из терминала.
# ------------------------------------------------------------------

import re
from decimal import Decimal
from pathlib import Path
import pandas as pd
import xlwings as xw

# ---------- Константы ------------------------------------------------------

EXCEL_PATH = Path(__file__).resolve().parents[1] / "Finmodel.xlsm"

SHEET_PLAN    = "ПланПродажОзон"
SHEET_COST    = "РасчётСебестоимости"
SHEET_SETTINGS= "Настройки"
SHEET_TARGET  = "РасчетЭкономикиОзон"          # куда выводим итоговую таблицу

MONTH_COLS    = [
    ("Мес.01", 1), ("Мес.02", 2), ("Мес.03", 3), ("Мес.04", 4),
    ("Мес.05", 5), ("Мес.06", 6), ("Мес.07", 7), ("Мес.08", 8),
    ("Мес.09", 9), ("Мес.10", 10), ("Мес.11", 11), ("Мес.12", 12),
]

# ---------- Вспомогательные функции ----------------------------------------

def _percent(value) -> Decimal:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return Decimal("0")
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    text = str(value).strip().replace(",", ".")
    if text.endswith("%"):
        text_num = text.rstrip("%").strip()
        return Decimal(text_num) / Decimal("100") if re.match(r"^\d*\.?\d+$", text_num) else Decimal("0")
    try:
        return Decimal(text) if re.match(r"^\d*\.?\d+$", text) else Decimal("0")
    except Exception:
        return Decimal("0")
    
def _get_workbook() -> tuple[xw.Book, xw.App, bool]:
    try:
        wb = xw.Book.caller()
        return wb, wb.app, False
    except Exception:
        app = xw.App(visible=False, add_book=False)
        wb  = app.books.open(EXCEL_PATH)
        return wb, app, True

def _load_settings(ws_settings) -> dict:
    used = ws_settings.used_range.value
    settings = {}
    for row in used:
        if not row or row[0] is None:
            continue
        param, val = str(row[0]).strip(), row[1]
        coef = _percent(val)
        if coef != 0:
            settings[param] = coef
    return settings

def _write_df_to_excel_table(ws_target, df: pd.DataFrame):
    header  = df.columns.tolist()
    n_rows  = len(df.index) + 1
    n_cols  = len(header)
    ws_target.clear()  # полностью чистим, включая форматы и таблицы
    target_rng = ws_target.range((1, 1)).resize(n_rows, n_cols)
    target_rng.value = [header] + df.values.tolist()
    try:
        list_objects_count = ws_target.api.ListObjects().Count
    except Exception:
        list_objects_count = 0
    if list_objects_count == 0:
        ws_target.api.ListObjects.Add(1, target_rng.api, None, 1).Name = "tbl_OzonEconomics"
    else:
        tbl = ws_target.api.ListObjects(1)
        tbl.Resize(target_rng.api)

def _format_numeric_columns(ws_target, columns, num_format="0"):
    tbl = ws_target.api.ListObjects(1)
    header_names = [col.Name for col in tbl.ListColumns]
    for col_name in columns:
        if col_name not in header_names:
            continue
        col_idx = header_names.index(col_name) + 1
        rng = tbl.ListColumns(col_idx).DataBodyRange
        rng.NumberFormat = num_format

def _format_table_rub(ws_target):
    rub_cols = [
        "План_шт", "ВыручкаБезСкидок_руб", "БаллыСкидки_руб", "ПрограммыПартнеров_руб", "Выручка_руб",
        "БазовоеВознаграждение_руб", "ВознаграждениеПослСкидок_руб", "УслугиДоставки_руб", "УслугиАгентов_руб",
        "УслугиFBO_руб", "Реклама_руб", "ДругиеУслуги_руб", "ИтогоРасходыМП_руб",
        "СебестоимостьПродаж_руб", "СебестоимостьБезНДС_руб",
        "СебестоимостьПродажНалог, ₽", "СебестоимостьПродажНалог_без_НДС, ₽",
        "ВаловаяПрибыль_Упр", "ВаловаяПрибыль_Налог",
    ]
    _format_numeric_columns(ws_target, rub_cols, num_format="0")

def _set_table_style_and_tab_color(ws_target):
    # Стиль таблицы: TableStyleMedium7 (Green, Medium 7)
    try:
        tbl = ws_target.api.ListObjects(1)
        tbl.TableStyle = "TableStyleMedium7"
    except Exception as e:
        print(f"Не удалось применить стиль таблицы: {e}")

    # Цвет ярлыка: #92D050 (RGB 146, 208, 80)
    try:
        ws_target.api.Tab.Color = 80 + 208*256 + 146*256*256
    except Exception as e:
        print(f"Не удалось задать цвет ярлыка: {e}")


def _add_totals_row(ws_target, df: pd.DataFrame) -> None:
    """Append subtotal row below the table."""
    from xlwings.utils import col_name

    last_row = len(df.index) + 1
    total_row = last_row + 1
    ws_target.range((total_row, 1)).value = "ИТОГО"
    ws_target.range((total_row, 1)).api.Font.Bold = True

    rub_cols = [
        "План_шт", "ВыручкаБезСкидок_руб", "БаллыСкидки_руб", "ПрограммыПартнеров_руб", "Выручка_руб",
        "БазовоеВознаграждение_руб", "ВознаграждениеПослСкидок_руб", "УслугиДоставки_руб", "УслугиАгентов_руб",
        "УслугиFBO_руб", "Реклама_руб", "ДругиеУслуги_руб", "ИтогоРасходыМП_руб",
        "СебестоимостьПродаж_руб", "СебестоимостьБезНДС_руб",
        "СебестоимостьПродажНалог, ₽", "СебестоимостьПродажНалог_без_НДС, ₽",
        "ВаловаяПрибыль_Упр", "ВаловаяПрибыль_Налог",
    ]

    for idx, col_name_hdr in enumerate(df.columns, start=1):
        if col_name_hdr not in rub_cols:
            continue
        letter = col_name(idx)
        if col_name_hdr == "СебестоимостьПродажНалог, ₽":
            formula = f"=SUBTOTAL(109,{letter}2:{letter}{last_row})"
        else:
            formula = f"=SUBTOTAL(9,{letter}2:{letter}{last_row})"
        cell = ws_target.range((total_row, idx))
        cell.formula = formula
        cell.api.Font.Bold = True
        cell.api.NumberFormat = "#,##0 ₽"


def _drop_totals(df: pd.DataFrame, check_cols=("Артикул_поставщика", "SKU", "Организация")) -> pd.DataFrame:
    """
    Удаляет строки, где в любом из check_cols встречается 'итого'
    (регистр и пробелы игнорируются).
    """
    if df.empty:
        return df

    def is_total(val):
        try:
            return str(val).strip().lower() == "итого"
        except Exception:
            return False

    mask = df.apply(
        lambda row: any(
            is_total(row[col]) for col in check_cols if col in df.columns
        ),
        axis=1,
    )
    return df.loc[~mask].copy()


def compute_ozon_economics_df(
    plan_df: pd.DataFrame, cost_df: pd.DataFrame, settings: dict
) -> pd.DataFrame:
    """Compute Ozon economics records from source DataFrames."""
    def get(key):
        return settings.get(key, Decimal("0"))

    tax_col_new = "СебестоимостьНалог_руб"
    tax_col_old = "Себестоимость_Налог, руб (новый)"
    tax_col_legacy = "СебестоимостьНалог"

    for col in ["СебестоимостьУпр", tax_col_new]:
        if col not in cost_df.columns:
            if col == tax_col_new and tax_col_old in cost_df.columns:
                cost_df[col] = cost_df[tax_col_old]
            elif col == tax_col_new and tax_col_legacy in cost_df.columns:
                cost_df[col] = cost_df[tax_col_legacy]
            else:
                cost_df[col] = 0

    cost_df = cost_df[[
        "Организация",
        "Артикул_поставщика",
        "Себестоимость_руб",
        "Себестоимость_без_НДС_руб",
        "СебестоимостьУпр",
        tax_col_new,
    ]]

    records: list[dict] = []
    for _, row in plan_df.iterrows():
        for col_name, month_num in MONTH_COLS:
            qty = row.get(col_name, 0) or 0
            if qty == 0:
                continue

            price = Decimal(str(row["Плановая цена"]))
            qty_dec = Decimal(str(qty))
            sales_brut = qty_dec * price

            bal_disc = sales_brut * get("Баллы за скидки")
            partner_prg = sales_brut * get("Программы партнеров")
            revenue = sales_brut - partner_prg - bal_disc

            base_comm = sales_brut * get("Вознаграждение Озон")

            def safe_dec(val, default="0"):
                try:
                    return Decimal(str(val))
                except Exception:
                    return Decimal(default)

            base_comm_099 = safe_dec(base_comm) * safe_dec("0.99")
            bal_disc_val = safe_dec(bal_disc)
            try:
                disc_limit = min(base_comm_099, bal_disc_val)
            except Exception:
                disc_limit = Decimal("0")
            try:
                comm_after = max(Decimal("0"), safe_dec(base_comm) - disc_limit)
            except Exception:
                comm_after = Decimal("0")

            serv_deliv = sales_brut * get("Услуги доставки")
            serv_agent = sales_brut * get("Услуги агентов")
            serv_fbo = sales_brut * get("Услуги FBO")
            reklama = sales_brut * get("Реклама")
            other_serv = sales_brut * get("Другие услуги")

            total_mp = comm_after + serv_deliv + serv_agent + serv_fbo + reklama + other_serv

            cs_row = cost_df[
                (cost_df["Организация"] == row["Организация"]) &
                (cost_df["Артикул_поставщика"] == row["Артикул_поставщика"])
            ]
            if cs_row.empty:
                cost_unit = Decimal("0")
                cost_unit_nds = Decimal("0")
                cost_mgmt = Decimal("0")
                cost_tax = Decimal("0")
            else:
                first = cs_row.iloc[0]
                cost_unit = Decimal(str(first["Себестоимость_руб"]))
                cost_unit_nds = Decimal(str(first["Себестоимость_без_НДС_руб"]))
                cost_mgmt = Decimal(str(first.get("СебестоимостьУпр", 0)))
                cost_tax = Decimal(
                    str(
                        first.get(
                            tax_col_new,
                            first.get(tax_col_old, first.get(tax_col_legacy, 0))
                        )
                    )
                )

            cogs = cost_unit * qty_dec
            cogs_no_vat = cost_unit_nds * qty_dec
            cogs_mgmt = cost_mgmt * qty_dec
            cogs_tax = cost_tax * qty_dec
            gp_mgmt = revenue - cogs_mgmt
            gp_tax = revenue - cogs_tax

            records.append(
                dict(
                    Месяц=month_num,
                    Организация=row["Организация"],
                    Артикул_поставщика=row["Артикул_поставщика"],
                    SKU=row["SKU"],
                    План_шт=qty,
                    ВыручкаБезСкидок_руб=float(sales_brut),
                    БаллыСкидки_руб=float(bal_disc),
                    ПрограммыПартнеров_руб=float(partner_prg),
                    Выручка_руб=float(revenue),
                    БазовоеВознаграждение_руб=float(base_comm),
                    ВознаграждениеПослСкидок_руб=float(comm_after),
                    УслугиДоставки_руб=float(serv_deliv),
                    УслугиАгентов_руб=float(serv_agent),
                    УслугиFBO_руб=float(serv_fbo),
                    Реклама_руб=float(reklama),
                    ДругиеУслуги_руб=float(other_serv),
                    ИтогоРасходыМП_руб=float(total_mp),
                    СебестоимостьПродаж_руб=float(cogs),
                    СебестоимостьБезНДС_руб=float(cogs_no_vat),
                    СебестоимостьНалог_руб=float(cogs_tax),
                    ВаловаяПрибыль_Упр=float(gp_mgmt),
                    ВаловаяПрибыль_Налог=float(gp_tax),
                )
            )

    result_df = (
        pd.DataFrame(records)
        .sort_values(["Месяц", "Организация", "Артикул_поставщика"])
        .reset_index(drop=True)
    )

    result_df = result_df.rename(columns={
        "СебестоимостьНалог_руб": "СебестоимостьПродажНалог, ₽",
    })
    result_df["СебестоимостьПродажНалог_без_НДС, ₽"] = (
        result_df["СебестоимостьПродажНалог, ₽"] / 1.2
    ).round(2)

    numeric_cols = [
        "План_шт",
        "ВыручкаБезСкидок_руб",
        "БаллыСкидки_руб",
        "ПрограммыПартнеров_руб",
        "Выручка_руб",
        "БазовоеВознаграждение_руб",
        "ВознаграждениеПослСкидок_руб",
        "УслугиДоставки_руб",
        "УслугиАгентов_руб",
        "УслугиFBO_руб",
        "Реклама_руб",
        "ДругиеУслуги_руб",
        "ИтогоРасходыМП_руб",
        "СебестоимостьПродаж_руб",
        "СебестоимостьБезНДС_руб",
        "СебестоимостьПродажНалог, ₽",
        "СебестоимостьПродажНалог_без_НДС, ₽",
        "ВаловаяПрибыль_Упр",
        "ВаловаяПрибыль_Налог",
    ]

    for col in numeric_cols:
        if col in result_df:
            result_df[col] = pd.to_numeric(result_df[col], errors="coerce")

    return result_df


def compute_wb_economics_df(plan_df: pd.DataFrame, cost_df: pd.DataFrame) -> pd.DataFrame:
    """Simplified computation of Wildberries economics for tests."""
    month_cols = [c for c in plan_df.columns if c.startswith("Мес.")]
    records: list[dict] = []
    for _, row in plan_df.iterrows():
        for col in month_cols:
            qty = row.get(col, 0) or 0
            if qty == 0:
                continue
            month_num = int(col.split(".")[-1])
            revenue = float(row.get("Выручка, ₽", 0))
            comm_rate = float(row.get("Комиссия WB %", 0))
            commission = revenue * comm_rate

            cs_row = cost_df[
                (cost_df["Организация"] == row["Организация"]) &
                (cost_df["Артикул_поставщика"] == row["Артикул_поставщика"])
            ]
            if cs_row.empty:
                c_unit = c_unit_wo = c_tax = c_tax_wo = 0.0
            else:
                first = cs_row.iloc[0]
                c_unit = float(first.get("Себестоимость_руб", 0))
                c_unit_wo = float(first.get("Себестоимость_без_НДС_руб", 0))
                c_tax = float(first.get("СебестоимостьНалог", c_unit_wo))
                c_tax_wo = float(first.get("СебестоимостьНалог_без_НДС", c_tax))

            cogs = c_unit * qty
            cogs_wo = c_unit_wo * qty
            cogs_tax = c_tax * qty
            cogs_tax_wo = c_tax_wo * qty
            gp_tax = revenue - cogs_tax
            ebit_tax = gp_tax - commission

            records.append(
                dict(
                    Организация=row["Организация"],
                    Артикул_WB=row.get("Артикул_WB"),
                    Артикул_поставщика=row["Артикул_поставщика"],
                    Предмет=row.get("Предмет", ""),
                    Месяц=month_num,
                    **{
                        "Кол-во, шт": qty,
                        "Выручка, ₽": revenue,
                        "Комиссия WB %": comm_rate,
                        "Комиссия WB, ₽": commission,
                        "Логистика, ₽": 0,
                        "Хранение, ₽": 0,
                        "Реклама, ₽": 0,
                        "Расходы МП, ₽": commission,
                        "СебестоимостьПродажРуб": cogs,
                        "СебестоимостьПродажБезНДС": cogs_wo,
                        "СебестоимостьПродажНалог, ₽": cogs_tax,
                        "СебестоимостьПродажНалог_без_НДС, ₽": cogs_tax_wo,
                        "ВаловаяПрибыль_Налог, ₽": gp_tax,
                        "EBITDA_Упр, ₽": revenue - commission - cogs,
                        "EBITDA_Налог, ₽": ebit_tax,
                        "ЧистаяПрибыль_Упр, ₽": revenue - commission - cogs,
                    },
                )
            )

    return pd.DataFrame(records)

# ---------- Основная логика -------------------------------------------------

def build_ozon_economics_table():
    wb, app, created = _get_workbook()
    try:
        ws_plan     = wb.sheets[SHEET_PLAN]
        ws_cost     = wb.sheets[SHEET_COST]
        ws_settings = wb.sheets[SHEET_SETTINGS]
        ws_target   = wb.sheets[SHEET_TARGET] if SHEET_TARGET in [s.name for s in wb.sheets] \
                      else wb.sheets.add(SHEET_TARGET, after=wb.sheets[SHEET_PLAN])
        try:
            if ws_target.index != 11:
                ws_target.api.Move(Before=wb.sheets[10].api)
                print("→ Лист 'РасчетЭкономикиОзон' перемещён на позицию 11")
        except Exception as e:
            print(f"⚠️ Не удалось переместить лист: {e}")
        cfg = _load_settings(ws_settings)

        plan_df = ws_plan.used_range.options(pd.DataFrame, header=1, index=False).value
        plan_df = _drop_totals(plan_df)  # <--- добавлено

        cost_df = ws_cost.used_range.options(pd.DataFrame, header=1, index=False).value
        cost_df = _drop_totals(cost_df)

        result_df = compute_ozon_economics_df(plan_df, cost_df, cfg)

        print(f"Строк для записи: {len(result_df)}")
        print("Колонки:", result_df.columns.tolist())
        print(result_df.head(3))

        _write_df_to_excel_table(ws_target, result_df)
        _format_table_rub(ws_target)
        _set_table_style_and_tab_color(ws_target)
        _add_totals_row(ws_target, result_df)


    finally:
        if created:
            wb.save()
            wb.close()
            app.quit()

# ---------- Точка входа -----------------------------------------------------

def main():
    build_ozon_economics_table()

if __name__ == "__main__":
    main()

