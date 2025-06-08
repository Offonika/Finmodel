# create_ozon_economics_table.py

import os
import xlwings as xw

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'Finmodel.xlsm')

SHEET_NAMES = {
    'plan_revenue':  'ПланВыручкиОзон',
    'plan_sales':    'ПланПродажОзон',
    'prices':        'ЦеныОзон',
    'avg_log':       'Показатели',
    'settings':      'Настройки',
    'cost':          'РасчётСебестоимости',
    'result':        'РасчетЭкономикиОзон'
}

MONTHS = [f'Мес.{str(i+1).zfill(2)}' for i in range(12)]

def get_workbook():
    try:
        wb = xw.Book.caller()
        app = None
        from_caller = True
    except Exception:
        app = xw.App(visible=False)
        wb = app.books.open(EXCEL_PATH)
        from_caller = False
    return wb, app, from_caller

def idx_from_header(header_row):
    return {str(h).strip(): i for i, h in enumerate(header_row)}

def to_num(val):
    try:
        if val is None or val == '':
            return 0
        return float(str(val).replace(',', '.').replace('₽', '').replace('%','').replace('\xa0','').replace(' ', ''))
    except Exception:
        return 0

def main():
    import time
    start = time.time()
    print('⏳ [START] Экономика Ozon')

    wb, app, from_caller = get_workbook()
    try:
        sh = {}
        for key, name in SHEET_NAMES.items():
            if name in [s.name for s in wb.sheets]:
                sh[key] = wb.sheets[name]
            else:
                if key == 'result':
                    sh[key] = wb.sheets.add(name)
                    print(f'➕ Лист "{name}" создан')
                else:
                    raise Exception(f'❌ Не найден лист: {name}')

        # Загружаем данные и индексы
        plan_revenue_data = sh['plan_revenue'].range(1,1).expand('table').value
        plan_sales_data   = sh['plan_sales'].range(1,1).expand('table').value
        prices_data       = sh['prices'].range(1,1).expand('table').value
        avg_log_data      = sh['avg_log'].range(1,1).expand('table').value
        settings_data     = sh['settings'].range(1,1).expand('table').value
        cost_data         = sh['cost'].range(1,1).expand('table').value

        idxPlanRevenue = idx_from_header(plan_revenue_data[0])
        idxPlanSales   = idx_from_header(plan_sales_data[0])
        idxPrices      = idx_from_header(prices_data[0])
        idxAvgLog      = idx_from_header(avg_log_data[0])
        idxCost        = idx_from_header(cost_data[0])

        # Индекс по артикулу — цены
        pricesIdx = {}
        for row in prices_data[1:]:
            art = str(row[idxPrices["Артикул"]]).strip()
            pricesIdx[art] = row

        # ДРР из Настройки
        drrVal = 0
        for r in settings_data:
            if str(r[0] or '').strip() == 'ДРР':
                drrVal = to_num(r[1]) / 100

        # Себестоимость — индекс по паре
        costIdxMap = {}
        for row in cost_data[1:]:
            key = f"{str(row[idxCost['Организация']]).strip()}||{str(row[idxCost['Артикул_поставщика']]).strip()}"
            costIdxMap[key] = row

        # Индексы для плана выручки и продаж
        planRevenueIdx = {}
        for row in plan_revenue_data[1:]:
            org = str(row[idxPlanRevenue['Организация']]).strip()
            art = str(row[idxPlanRevenue['Артикул_поставщика']]).strip()
            for m in range(1, 13):
                mName = f'Мес.{str(m).zfill(2)}'
                rev = to_num(row[idxPlanRevenue.get(mName, -1)])
                if rev:
                    planRevenueIdx[f'{org}||{art}||{mName}'] = rev

        planSalesIdx = {}
        for row in plan_sales_data[1:]:
            org = str(row[idxPlanSales['Организация']]).strip()
            art = str(row[idxPlanSales['Артикул_поставщика']]).strip()
            if not org or not art or org.lower() == "итого" or art.lower() == "итого" or org == 'None' or art == 'None':
                continue
            for m in range(1, 13):
                mName = f'Мес.{str(m).zfill(2)}'
                qty = to_num(row[idxPlanSales.get(mName, -1)])
                if qty:
                    planSalesIdx[f'{org}||{art}||{mName}'] = qty

        # Средние затраты
        avgLogRow = avg_log_data[1] if len(avg_log_data) > 1 else None

        HEADER = [
            "Месяц", "Организация", "Артикул_поставщика", "Кол-во, шт", "Выручка, ₽",
            "Комиссия %", "Комиссия, ₽", "Оплата эквайринга, ₽",
            "Логистика, ₽", "Обработка отправления, ₽", "Магистраль, ₽", "Последняя миля, ₽",
            "Обратная магистраль, ₽", "Обработка возврата, ₽", "Обратная логистика, ₽",
            "Реклама, ₽", "Расходы МП, ₽", "СебестоимостьПродажРуб", "Себестоимость_без_НДС_руб"
        ]
        result = []

        # Основной цикл
        for row in plan_sales_data[1:]:
            org = str(row[idxPlanSales['Организация']]).strip()
            art = str(row[idxPlanSales['Артикул_поставщика']]).strip()
            if not org or not art or org.lower() in ("итого", "none") or art.lower() in ("итого", "none"):
                continue
            for m in range(1, 13):
                mName = f'Мес.{str(m).zfill(2)}'
                qty = to_num(row[idxPlanSales.get(mName, -1)])
                if not qty:
                    continue
                revenue = planRevenueIdx.get(f'{org}||{art}||{mName}', 0)
                if not revenue:
                    continue
                priceRow = pricesIdx.get(art, [])
                commPerc = to_num(priceRow[idxPrices.get("FBO: % продажи", -1)]) / 100 if priceRow else 0
                commRub  = revenue * commPerc

                ekvPerc = to_num(avgLogRow[idxAvgLog.get("Эквайринг, %", -1)]) / 100 if avgLogRow else 0
                ekvRub  = revenue * ekvPerc

                def avgLogVal(name):
                    return to_num(avgLogRow[idxAvgLog.get(name, -1)]) * qty if avgLogRow else 0

                logistika = avgLogVal("Логистика, ₽")
                otpravka  = avgLogVal("Обработка отправления, ₽")
                magistral = avgLogVal("Магистраль, ₽")
                lastMile  = avgLogVal("Последняя миля, ₽")
                obrMag    = avgLogVal("Обратная магистраль, ₽")
                obrReturn = avgLogVal("Обработка возврата, ₽")
                obrLog    = avgLogVal("Обратная логистика, ₽")

                reklRub = revenue * drrVal
                mpCosts = commRub + ekvRub + logistika + otpravka + magistral + lastMile + obrMag + obrReturn + obrLog + reklRub

                costRow = costIdxMap.get(f'{org}||{art}', [])
                cogsRub   = to_num(costRow[idxCost.get("Себестоимость_руб", -1)]) * qty if costRow else 0
                cogsNoVat = to_num(costRow[idxCost.get("Себестоимость_без_НДС_руб", -1)]) * qty if costRow else 0

                result.append([
                    mName, org, art, qty, revenue,
                    commPerc * 100, round(commRub), round(ekvRub),
                    round(logistika), round(otpravka), round(magistral), round(lastMile),
                    round(obrMag), round(obrReturn), round(obrLog),
                    round(reklRub), round(mpCosts), round(cogsRub), round(cogsNoVat)
                ])


        # --- Запись и оформление Excel Table ---
        sheet = sh['result']
        sheet.clear()
        sheet.range((1, 1), (1, len(HEADER))).value = HEADER
        if result:
            sheet.range((2, 1), (len(result) + 1, len(HEADER))).value = result

        # Удалить старую таблицу
        try:
            for tbl in sheet.api.ListObjects:
                if tbl.Name == 'OzonEconomicsTable':
                    tbl.Delete()
        except Exception:
            pass

        last_row = len(result) + 1
        last_col = len(HEADER)
        table_range = sheet.range((1, 1), (last_row, last_col))
        lo = sheet.api.ListObjects.Add(
            SourceType=1,
            Source=table_range.api,
            XlListObjectHasHeaders=1
        )
        lo.Name = 'OzonEconomicsTable'
        lo.TableStyle = 'TableStyleMedium7'

        # Форматы
        # Комиссия % (6-й столбец)
        sheet.range((2, 6), (last_row, 6)).api.NumberFormat = '0.00"%"'
        # Рубли (5,7-18)
        for c in [5, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]:
            sheet.range((2, c), (last_row, c)).api.NumberFormat = '0 ₽'
        # Кол-во, шт (4)
        sheet.range((2, 4), (last_row, 4)).api.NumberFormat = '0'

        sheet.api.Tab.Color = 5296274  # зелёный
        sheet.api.Move(Before=wb.sheets[3].api)  # сделать вторым листом
        sheet.autofit()
        sheet.api.Rows.AutoFit()
# --- Закрепить только первую строку (шапку)
        sheet.api.Activate()
        sheet.api.Application.ActiveWindow.FreezePanes = False
        sheet.range("A2").select()
        sheet.api.Application.ActiveWindow.FreezePanes = True

        wb.save()
        print(f'🏁 [FINISH] {len(result)} строк, {round(time.time() - start, 1)} сек.')
    finally:
        if not from_caller and wb is not None:
            wb.close()
        if app is not None:
            app.quit()
            del app

if __name__ == '__main__':
    main()
