import logging
import sys
from pathlib import Path
import xlwings as xw

IS_EXE = getattr(sys, "frozen", False)
BASE_DIR = Path(sys.executable).resolve().parent if IS_EXE else Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent if IS_EXE else BASE_DIR

EXCEL_PATH = PROJECT_DIR / "Finmodel.xlsm"
LOG_DIR = PROJECT_DIR / "log"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / "template_xlwings.log"

logging.basicConfig(
    filename=str(LOG_FILE),
    filemode="w",
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def get_workbook():
    """Return (wb, app). ``app`` is ``None`` when called from Excel."""
    try:
        wb = xw.Book.caller()
        app = None
        logging.info("→ Запуск из Excel")
    except Exception:
        if not EXCEL_PATH.exists():
            logging.error("Workbook not found: %s", EXCEL_PATH)
            raise FileNotFoundError(f"Workbook not found: {EXCEL_PATH}")
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(EXCEL_PATH)
        logging.info("→ Открыт файл: %s", EXCEL_PATH)
    return wb, app


def main():
    wb, app = get_workbook()
    logging.info("Template script started")
    # TODO: add your code here
    if app:
        wb.close()
        app.quit()
        logging.info("Excel closed")


if __name__ == "__main__":
    main()
