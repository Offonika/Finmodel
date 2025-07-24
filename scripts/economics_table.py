# economics_table.py
# ------------------------------------------------------------------
# Создание таблицы «РасчетЭкономикиОзон» по данным листов
# «ПланПродажОзон», «РасчётСебестоимости» и «Настройки».
# Скрипт работает как из Excel-макроса (RunPython), так и из терминала.
# ------------------------------------------------------------------

import os
import re
from decimal import Decimal
import pandas as pd
import xlwings as xw

# ---------- Константы ------------------------------------------------------

BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH    = os.path.join(BASE_DIR, "excel", "Finmodel.xlsm")

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
        get = lambda key: cfg.get(key, Decimal("0"))

        plan_df = ws_plan.used_range.options(pd.DataFrame, header=1, index=False).value
        plan_df = _drop_totals(plan_df)  # <--- добавлено

        cost_df = ws_cost.used_range.options(pd.DataFrame, header=1, index=False).value
        cost_df = _drop_totals(cost_df)  # <--- добавлено

        tax_col_new = "Себестоимость_Налог, руб (новый)"
        tax_col_old = "СебестоимостьНалог"

        for col in ["СебестоимостьУпр", tax_col_new]:
            if col not in cost_df.columns:
                if col == tax_col_new and tax_col_old in cost_df.columns:
                    cost_df[col] = cost_df[tax_col_old]
                else:
                    cost_df[col] = 0

        cost_df = cost_df[[
            "Организация", "Артикул_поставщика",
            "Себестоимость_руб", "Себестоимость_без_НДС_руб",
            "СебестоимостьУпр", tax_col_new,
        ]]

        records = []
        for _, row in plan_df.iterrows():
            for col_name, month_num in MONTH_COLS:
                qty = row.get(col_name, 0) or 0
                if qty == 0:
                    continue
                price   = Decimal(str(row["Плановая цена"]))
                qty_dec = Decimal(str(qty))
                sales_brut  = qty_dec * price

                bal_disc    = sales_brut * get("Баллы за скидки")
                partner_prg = sales_brut * get("Программы партнеров")
                revenue     = sales_brut - partner_prg - bal_disc


                base_comm   = sales_brut * get("Вознаграждение Озон")

                def safe_decimal(val, default="0"):
                    try:
                        return Decimal(str(val))
                    except Exception:
                        return Decimal(default)

                base_comm_099 = safe_decimal(base_comm) * safe_decimal("0.99")
                bal_disc_val  = safe_decimal(bal_disc)

                try:
                    disc_limit = min(base_comm_099, bal_disc_val)
                except Exception:
                    disc_limit = Decimal("0")

                try:
                    comm_after = max(Decimal("0"), safe_decimal(base_comm) - disc_limit)
                except Exception:
                    comm_after = Decimal("0")

                serv_deliv  = sales_brut * get("Услуги доставки")
                serv_agent  = sales_brut * get("Услуги агентов")
                serv_fbo    = sales_brut * get("Услуги FBO")
                reklama     = sales_brut * get("Реклама")
                other_serv  = sales_brut * get("Другие услуги")

                total_mp    = comm_after + serv_deliv + serv_agent + serv_fbo + reklama + other_serv

                cs_row = cost_df[
                    (cost_df["Организация"] == row["Организация"]) &
                    (cost_df["Артикул_поставщика"] == row["Артикул_поставщика"])
                ]
                if cs_row.empty:
                    cost_unit     = Decimal("0")
                    cost_unit_nds = Decimal("0")
                    cost_mgmt     = Decimal("0")
                    cost_tax      = Decimal("0")
                else:
                    first = cs_row.iloc[0]
                    cost_unit     = Decimal(str(first["Себестоимость_руб"]))
                    cost_unit_nds = Decimal(str(first["Себестоимость_без_НДС_руб"]))
                    cost_mgmt     = Decimal(str(first.get("СебестоимостьУпр", 0)))
                    cost_tax      = Decimal(str(first.get(tax_col_new, first.get(tax_col_old, 0))))

                cogs        = cost_unit     * qty_dec
                cogs_no_vat = cost_unit_nds * qty_dec
                cogs_mgmt   = cost_mgmt     * qty_dec
                cogs_tax    = cost_tax      * qty_dec
                gp_mgmt     = revenue - cogs_mgmt
                gp_tax      = revenue - cogs_tax

                records.append(dict(
                    Месяц             = month_num,
                    Организация       = row["Организация"],
                    Артикул_поставщика= row["Артикул_поставщика"],
                    SKU               = row["SKU"],
                    План_шт           = qty,
                    ВыручкаБезСкидок_руб      = float(sales_brut),
                    БаллыСкидки_руб           = float(bal_disc),
                    ПрограммыПартнеров_руб    = float(partner_prg),
                    Выручка_руб               = float(revenue),
                    БазовоеВознаграждение_руб = float(base_comm),
                    ВознаграждениеПослСкидок_руб = float(comm_after),
                    УслугиДоставки_руб        = float(serv_deliv),
                    УслугиАгентов_руб         = float(serv_agent),
                    УслугиFBO_руб             = float(serv_fbo),
                    Реклама_руб               = float(reklama),
                    ДругиеУслуги_руб          = float(other_serv),
                    ИтогоРасходыМП_руб        = float(total_mp),
                    СебестоимостьПродаж_руб   = float(cogs),
                    СебестоимостьБезНДС_руб   = float(cogs_no_vat),
                    ВаловаяПрибыль_Упр        = float(gp_mgmt),
                    ВаловаяПрибыль_Налог      = float(gp_tax),
                ))

        result_df = (
            pd.DataFrame(records)
            .sort_values(["Месяц", "Организация", "Артикул_поставщика"])
            .reset_index(drop=True)
        )

        # === Привести нужные столбцы к числам (float) ===
        numeric_cols = [
            "План_шт", "ВыручкаБезСкидок_руб", "БаллыСкидки_руб", "ПрограммыПартнеров_руб", "Выручка_руб",
            "БазовоеВознаграждение_руб", "ВознаграждениеПослСкидок_руб", "УслугиДоставки_руб", "УслугиАгентов_руб",
            "УслугиFBO_руб", "Реклама_руб", "ДругиеУслуги_руб", "ИтогоРасходыМП_руб",
            "СебестоимостьПродаж_руб", "СебестоимостьБезНДС_руб",
            "ВаловаяПрибыль_Упр", "ВаловаяПрибыль_Налог",
        ]
        for col in numeric_cols:
            if col in result_df:
                result_df[col] = pd.to_numeric(result_df[col], errors='coerce')

        # --- Диагностика ---
        print(f"Строк для записи: {len(result_df)}")
        print("Колонки:", result_df.columns.tolist())
        print(result_df.head(3))

        _write_df_to_excel_table(ws_target, result_df)
        _format_table_rub(ws_target)
        _set_table_style_and_tab_color(ws_target)


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

