# main.py
# main_service_charges.py  – агрегация «НачисленияУслугОзон»
# ---------------------------------------------------------
import os
import argparse
import xlwings as xw

from file_loader  import load_files
from finmodel.aggregator   import aggregate_data
from excel_writer import write_to_excel, write_df_to_excel_table


# ─── 1. Базовые относительные пути ────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))          # …/Finmodel/finmodel
ROOT_DIR   = os.path.abspath(os.path.join(BASE_DIR, os.pardir))  # …/Finmodel

DEFAULT_ORG = os.path.join(
    ROOT_DIR, "НачисленияУслугОзон", "ИП Закирова Р.Х"
)
EXCEL_BOOK = os.path.join(ROOT_DIR, "Finmodel.xlsm")

SHEET_NAME = "НачисленияУслугОзон"
TABLE_NAME = "НачисленияУслугОзонTable"


# ─── 2. Универсальный доступ к книге Excel ────────────────────────────────────
def get_workbook(excel_path: str = EXCEL_BOOK):
    """
    Возвращает (wb, app).
    • Если вызов из макроса RunPython – wb уже открыт, app=None.
    • Если запуск из терминала – открываем книгу сами в скрытом Excel.
    """
    try:
        wb  = xw.Book.caller()      # Excel → Python
        app = None
        print("→ Запуск из Excel-макроса")
    except Exception:
        app = xw.App(visible=False)
        wb  = app.books.open(excel_path)
        print("→ Консоль-режим:", excel_path)
    return wb, app


# ─── 3. Основная логика ───────────────────────────────────────────────────────
def run(org_folder: str = DEFAULT_ORG, excel_path: str = EXCEL_BOOK):
    files_df   = load_files(org_folder)
    result_df  = aggregate_data(files_df)

    wb, app = get_workbook(excel_path)

    # 1) просто вставляем DataFrame на лист
    write_to_excel(result_df, wb.fullname, sheet_name=SHEET_NAME)

    # 2) превращаем диапазон в умную таблицу
    write_df_to_excel_table(result_df, wb.fullname,
                            sheet_name=SHEET_NAME,
                            table_name=TABLE_NAME)

    if app is not None:            # закрываем, только если сами открывали
        wb.save()
        app.quit()


# ─── 4. CLI-интерфейс для терминала ───────────────────────────────────────────
def main():
    P = argparse.ArgumentParser(
        description="Агрегация начислений услуг Ozon → Finmodel.xlsm"
    )
    P.add_argument("-d", "--dir",  default=DEFAULT_ORG,
                   help="Папка с исходными отчётами")
    P.add_argument("-f", "--file", default=EXCEL_BOOK,
                   help="Путь к Finmodel.xlsm")
    args = P.parse_args()

    run(args.dir, args.file)


if __name__ == "__main__":
    main()
