# ozon_transactions_to_excel.py
import os
import xlwings as xw
import requests

# ==== НАСТРОЙКИ ====
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Ozon_Transactions.xlsx')
SHEET_NAME = 'Транзакции'

CLIENT_ID = '157232'
API_KEY = '7e9139cb-7dc3-4174-a6a1-41f839080bb2'
API_URL = "https://api-seller.ozon.ru/v3/finance/transaction/list"

# ==== ЗАГОЛОВКИ (русские) ====
HEADERS = [
    "ID операции",
    "Тип операции",
    "Дата операции",
    "Название типа операции",
    "Стоимость доставки",
    "Плата за возврат",
    "Начислено за продажу",
    "Комиссия за продажу",
    "Сумма операции",
    "Тип начисления",
    "Схема доставки",
    "Дата заказа",
    "Номер отправления",
    "ID склада"
]

def fetch_transactions(from_date, to_date):
    page = 1
    page_size = 1000
    all_ops = []
    while True:
        body = {
            "filter": {
                "date": {"from": from_date, "to": to_date},
                "operation_type": [],
                "posting_number": "",
                "transaction_type": "all"
            },
            "page": page,
            "page_size": page_size
        }
        resp = requests.post(
            API_URL,
            headers={
                "Client-Id": CLIENT_ID,
                "Api-Key": API_KEY,
                "Content-Type": "application/json"
            },
            json=body
        )
        resp.raise_for_status()
        data = resp.json()
        ops = data.get("result", {}).get("operations", [])
        all_ops.extend(ops)
        print(f"Загружена страница {page}, операций: {len(ops)}")
        if page >= data["result"]["page_count"] or not ops:
            break
        page += 1
    return all_ops

def prepare_rows(ops):
    """Конвертируем данные в строки для Excel с русскими заголовками"""
    rows = []
    for op in ops:
        rows.append([
            op.get("operation_id"),
            op.get("operation_type"),
            op.get("operation_date"),
            op.get("operation_type_name"),
            op.get("delivery_charge"),
            op.get("return_delivery_charge"),
            op.get("accruals_for_sale"),
            op.get("sale_commission"),
            op.get("amount"),
            op.get("type"),
            op.get("posting", {}).get("delivery_schema"),
            op.get("posting", {}).get("order_date"),
            op.get("posting", {}).get("posting_number"),
            op.get("posting", {}).get("warehouse_id")
        ])
    return rows

def write_to_excel(rows):
    # Если файла нет — создаём новый, если есть — обновляем лист
    app = xw.App(visible=False)
    if not os.path.exists(EXCEL_PATH):
        wb = app.books.add()
        wb.save(EXCEL_PATH)
    else:
        wb = app.books.open(EXCEL_PATH)
    # Удалить лист если уже был, чтобы не плодить копии
    if SHEET_NAME in [s.name for s in wb.sheets]:
        wb.sheets[SHEET_NAME].delete()
    ws = wb.sheets.add(SHEET_NAME, before=wb.sheets[0])
    ws.range("A1").value = [HEADERS] + rows
    # Форматируем шапку
    ws.range("A1:N1").api.Font.Bold = True
    ws.autofit()
    wb.save()
    app.quit()
    print(f"✅ Данные записаны в {EXCEL_PATH} (лист '{SHEET_NAME}')")

def main():
    # Пример периода: 1 ноября 2021 — 2 ноября 2021
    from_date = "2025-03-01T00:00:00.000Z"
    to_date   = "2025-03-31T23:59:59.000Z"
    print("Запрашиваем данные Ozon...")
    ops = fetch_transactions(from_date, to_date)
    print(f"Всего операций: {len(ops)}")
    rows = prepare_rows(ops)
    write_to_excel(rows)

if __name__ == "__main__":
    main()
