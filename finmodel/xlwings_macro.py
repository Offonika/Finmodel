# xlwings_macro.py  ─ запускается из Excel как python_module
# ----------------------------------------------------------
import os
import sys

# 1) гарантируем наличие корня репозитория в sys.path
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)          # теперь import finmodel.* всегда работает

# 2) обычные импорты
import xlwings as xw
from finmodel.main import main                               # ← было finmodel.main (оставляем)
from finmodel.wb_prices import load_wb_prices_by_size_xlwings  # ← добавили префикс

# 3) «обёртки» для вызова из VBA
def run_aggregation():
    try:
        main()
        xw.apps.active.api.MsgBox("Обработка завершена", 0, "Готово")
    except Exception as e:
        xw.apps.active.api.MsgBox(f"Ошибка: {str(e)}", 0, "Ошибка")

def run_wb_prices_by_size():
    try:
        load_wb_prices_by_size_xlwings()
        xw.apps.active.api.MsgBox("Загрузка цен по размерам завершена!", 0, "Готово")
    except Exception as e:
        xw.apps.active.api.MsgBox(f"Ошибка: {str(e)}", 0, "Ошибка")
