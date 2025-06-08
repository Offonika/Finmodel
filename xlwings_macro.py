import xlwings as xw
from main import main
from wb_prices import load_wb_prices_by_size_xlwings


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
