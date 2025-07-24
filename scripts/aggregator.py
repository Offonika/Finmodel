# aggregator.py

import pandas as pd
import xlwings as xw
from collections import defaultdict
import logging
from datetime import datetime
import os

from scripts.style_utils import format_table

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, 'excel', 'Finmodel.xlsm')
SHEET_NAME = 'НачисленияУслугОзон'

def get_workbook():
    try:
        wb = xw.Book.caller()
        return wb, wb.app, False
    except Exception:
        app = xw.App(visible=False, add_book=False)
        wb  = app.books.open(EXCEL_PATH)
        return wb, app, True

def write_to_excel(df: pd.DataFrame):
    wb, app, created = get_workbook()
    try:
        if SHEET_NAME not in [s.name for s in wb.sheets]:
            wb.sheets.add(SHEET_NAME)
        sht = wb.sheets[SHEET_NAME]
        sht.clear_contents()
        sht.range("A1").value = df.columns.tolist()
        sht.range("A2").value = df.values.tolist()

        # --- Оформление таблицы и шапки ---
        last_row = df.shape[0] + 1  # +1 для шапки
        last_col = df.shape[1]
        table_range = sht.range((1, 1), (last_row, last_col))
        format_table(sht, table_range, "ServiceChargesTable")
        sht.api.Application.ActiveWindow.SplitRow = 1
        sht.api.Application.ActiveWindow.FreezePanes = True

        # --- Цвет ярлыка #00FFC0 (бирюза в BGR) ---
        try:
            sht.api.Tab.Color = 0xC0FF00
            print("→ Цвет ярлыка #00FFC0 установлен")
        except Exception as e:
            print(f"⚠️ Не удалось установить цвет ярлыка: {e}")

        # --- Переместить лист на позицию 12 ---
        try:
            if sht.index != 12:
                sht.api.Move(Before=sht.book.sheets[11].api)
                print("→ Лист перемещён на позицию 12")
        except Exception as e:
            print(f"⚠️ Не удалось переместить лист: {e}")

        wb.save()
    finally:
        if created:
            wb.close()
            app.quit()

def setup_logger(log_name='aggregator'):
    log_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(log_dir, f"{log_name}_{datetime.now():%Y%m%d}.log")
    logger = logging.getLogger(log_name)
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(logging.Formatter('%(asctime)s %(levelname)s: %(message)s'))
    if not logger.handlers:
        logger.addHandler(fh)
    return logger

HEADER = [
    'Месяц','Организация','SKU','Артикул','Название','ПроданоШт','ВыручкаБезСкидок','Комиссия',
    'Начисление за хранение/утилизацию возвратов','Оплата эквайринга','Сборка заказа','Обработка отправления',
    'Магистраль','Последняя миля','Обратная магистраль','Обработка возврата','Логистика',
    'Обратная логистика','ИтогЛогистика','Начисление за гибкий график выплат','ЗатратыНа продвижение',
    'ДругиеУслуги','УслугиПартнеров','Продвижение бренда','Баллы за отзывы','Вывод в топ','Подписка Premium',
    'Временное размещение товара в СЦ/ПВЗ','Услуга досрочной выплаты',
    'Услуга за обработку операционных ошибок продавца: поздняя отгрузка',
    'Услуга за обработку операционных ошибок продавца: поздняя отгрузка - отмена начисления',
    'Утилизация товара: Вы не забрали в срок',
    'Услуга размещения товаров на складе',
    'Услуги FBO',
    'Компенсации',
    'Прибыль'
]

def get(row, name):
    try:
        return float(row.get(name, 0) or 0)
    except Exception:
        return 0

def aggregate_data(df: pd.DataFrame) -> pd.DataFrame:
    logger = setup_logger('aggregator')
    logger.info('Старт агрегации данных')
    df = df.copy()
    df['Месяц'] = pd.to_datetime(df['Дата начисления'], errors='coerce').dt.to_period('M').astype(str)
    df['Организация'] = df['organization'] if 'organization' in df.columns else ''

    # --- Маппинг для новых вычислений ---
    fbo_types = [
        "Доставка товаров на склад Ozon (кросс-докинг)",
        "Обработка товара в составе грузоместа на FBO",
        "Услуга по бронированию места и персонала для поставки с неполным составом",
        "Услуга по бронированию места и персонала для поставки с неполным составом в составе ГМ",
        "Услуга по обработке опознанных излишков",
        "Услуга по обработке опознанных излишков в составе ГМ",
        "Услуга размещения товаров на складе"
    ]
    partner_types = [
        "Звёздные товары",
        "Услуги международной доставки",
        "Упаковка товара партнерами"
    ]
    promo_types = [
        "Вывод в топ",
        "Подписка Premium",
        "Продвижение бренда",
        "Баллы за отзывы",
        "Бонусы продавца",
        "Бонусы продавца - рассылка",
        "Подписка Premium Plus",
        "Продвижение в поиске",
        "Трафареты",
        "Абонентское обслуживание по продвижению товаров"
    ]
    other_services_types = [
        "Временное размещение товара в СЦ/ПВЗ",
        "Начисление за гибкий график выплат",
        "Услуга досрочной выплаты",
        "Услуга за обработку опер. ошибок: отмена",
        "Услуга за обработку опер. ошибок: поздняя отгрузка",
        "Утилизация",
        "Утилизация товара: Вы не забрали в срок",
        "Утилизация товара: Повреждённые из-за упаковки",
        "Утилизация товара: Повреждённые, были у покупателя",
        "Утилизация товара: Прочее"
    ]
    compensation_types = [
        "Декомпенсации и возвращение товаров на сток",
        "Потеря по вине Ozon в логистике",
        "Потеря по вине Ozon на складе",
        "Брак по вине Ozon на складе"
    ]

    sales_types = {
        "Доставка покупателю": "sale",
        "Получение возврата, отмены, невыкупа от покупателя": "return"
    }
    revenue_types = [
        "Доставка покупателю",
        "Доставка покупателю — отмена начисления",
        "Получение возврата, отмены, невыкупа от покупателя"
    ]

    groups = defaultdict(lambda: defaultdict(float))
    logger.info(f'Строк для обработки: {len(df)}')

    for idx, row in df.iterrows():
        if idx % 10000 == 0 and idx > 0:
            logger.info(f"Обработано строк: {idx}")
        key = (
            row.get('Месяц', ''),
            row.get('Организация', ''),
            str(row.get('SKU', '')).strip() or '',
            str(row.get('Артикул', '')).strip() or '',
            '' if (pd.isna(row.get('SKU')) and pd.isna(row.get('Артикул'))) else str(row.get('Название товара или услуги', '')).strip()
        )

        accr_type = str(row.get('Тип начисления', '')).strip()

        # Количество продаж и возвратов
        if accr_type in sales_types:
            t = sales_types[accr_type]
            groups[key][f'qty_{t}'] += float(row.get('Количество', 0) or 0)
        if accr_type in revenue_types:
            groups[key]['revenue'] += float(row.get('За продажу или возврат до вычета комиссий и услуг', 0) or 0)
        if accr_type in ["Доставка покупателю", "Доставка покупателю — отмена начисления"]:
            groups[key]['com_sale'] += float(row.get('Комиссия за продажу', 0) or 0)
        if accr_type == "Получение возврата, отмены, невыкупа от покупателя":
            groups[key]['com_return'] += float(row.get('Комиссия за продажу', 0) or 0)

        # Основные затраты и услуги
        groups[key]['Обработка возврата'] += float(row.get('Обработка возврата', 0) or 0)
        groups[key]['Последняя миля'] += float(row.get('Последняя миля (разбивается по товарам пропорционально доле цены товара в сумме отправления)', 0) or 0)
        groups[key]['Обратная магистраль'] += float(row.get('Обратная магистраль', 0) or 0)
        if accr_type == "Оплата эквайринга":
            groups[key]['Оплата эквайринга'] += float(row.get('Итого', 0) or 0)

        # Для FBO, Партнеров, Продвижения, Прочих услуг — агрегируем отдельно
        if accr_type in fbo_types:
            groups[key]['Услуги FBO'] += float(row.get('Итого', 0) or 0)
        if accr_type in partner_types:
            groups[key]['Партнерские услуги'] += float(row.get('Итого', 0) or 0)
        if accr_type in promo_types:
            groups[key]['ЗатратыНа продвижение'] += float(row.get('Итого', 0) or 0)
        if accr_type in other_services_types:
            groups[key]['ДругиеУслуги'] += float(row.get('Итого', 0) or 0)
        if accr_type in compensation_types:
            groups[key]['Компенсации'] += float(row.get('Итого', 0) or 0)

        # Для обратной совместимости (старая структура)
        groups[key]['Начисление за хранение/утилизацию возвратов'] += float(row.get('Начисление за хранение/утилизацию возвратов', 0) or 0)
        groups[key]['Сборка заказа'] += float(row.get('Сборка заказа', 0) or 0)
        groups[key]['Обработка отправления'] += float(row.get('Обработка отправления (Drop-off/Pick-up) (разбивается по товарам пропорционально количеству в отправлении)', 0) or 0)
        groups[key]['Магистраль'] += float(row.get('Магистраль', 0) or 0)
        groups[key]['Логистика'] += float(row.get('Логистика', 0) or 0)
        groups[key]['Обратная логистика'] += float(row.get('Обратная логистика', 0) or 0)
        # Старые отдельные начисления (нужны для некоторых полей)
        groups[key]['Услуга за обработку операционных ошибок продавца: поздняя отгрузка'] += float(row.get('Итого', 0) or 0) if accr_type == 'Услуга за обработку операционных ошибок продавца: поздняя отгрузка' else 0
        groups[key]['Услуга за обработку операционных ошибок продавца: поздняя отгрузка - отмена начисления'] += float(row.get('Итого', 0) or 0) if accr_type == 'Услуга за обработку операционных ошибок продавца: поздняя отгрузка - отмена начисления' else 0
        groups[key]['Утилизация товара: Вы не забрали в срок'] += float(row.get('Итого', 0) or 0) if accr_type == 'Утилизация товара: Вы не забрали в срок' else 0

        if accr_type == "Продвижение бренда":
            groups[key]['Продвижение бренда'] += float(row.get('Итого', 0) or 0)
        if accr_type == "Баллы за отзывы":
            groups[key]['Баллы за отзывы'] += float(row.get('Итого', 0) or 0)
        if accr_type == "Вывод в топ":
            groups[key]['Вывод в топ'] += float(row.get('Итого', 0) or 0)
        if accr_type == "Подписка Premium":
            groups[key]['Подписка Premium'] += float(row.get('Итого', 0) or 0)

    result = []
    for key, vals in groups.items():
        sold_qty = vals.get('qty_sale', 0) - vals.get('qty_return', 0)
        revenue = vals.get('revenue', 0)
        commission = vals.get('com_sale', 0) + vals.get('com_return', 0)
        posled_milya = vals.get('Последняя миля', 0)
        obr_magistral = vals.get('Обратная магистраль', 0)
        oplata_ekv = vals.get('Оплата эквайринга', 0)

        # --- Новая логика по твоим правилам ---
        uslugi_fbo = vals.get('Услуги FBO', 0)
        uslugi_partnerov = posled_milya + obr_magistral - oplata_ekv + vals.get('Партнерские услуги', 0)
        zatraty_na_prodvizhenie = vals.get('ЗатратыНа продвижение', 0)
        drugie_uslugi = vals.get('ДругиеУслуги', 0)

        # Итоговая логистика (старое)
        vals['ИтогЛогистика'] = vals.get('Логистика', 0) + vals.get('Обратная логистика', 0)

        profit = (
            revenue
            + commission
            + vals.get('ИтогЛогистика', 0)
            + drugie_uslugi
            + uslugi_partnerov
            + uslugi_fbo
            + vals.get('Компенсации', 0)
        )

        row = list(key)
        row += [
            sold_qty,
            revenue,
            commission
        ]

        # Динамическое заполнение остальных полей по HEADER
        for col in HEADER[8:-1]:  # -1 чтобы прибыль добавить в конце
            if col == "Последняя миля":
                row.append(posled_milya)
            elif col == "Обратная магистраль":
                row.append(obr_magistral)
            elif col == "Оплата эквайринга":
                row.append(oplata_ekv)
            elif col == "Услуги FBO":
                row.append(uslugi_fbo)
            elif col == "УслугиПартнеров":
                row.append(uslugi_partnerov)
            elif col == "ЗатратыНа продвижение":
                row.append(zatraty_na_prodvizhenie)
            elif col == "ДругиеУслуги":
                row.append(drugie_uslugi)
            else:
                row.append(vals.get(col, 0))
        row.append(profit)
        result.append(row)

    result_df = pd.DataFrame(result, columns=HEADER)
    logger.info(f'Агрегация завершена. Получено строк: {len(result_df)}')
    return result_df


def main():
    logger = setup_logger('aggregator')
    orgs_dir = os.path.join(BASE_DIR, 'НачисленияУслугОзон')
    if not os.path.exists(orgs_dir):
        logger.error(f"Нет папки: {orgs_dir}")
        print(f"❌ Нет папки: {orgs_dir}")
        return

    # Ищем все подкаталоги (по организациям)
    org_folders = [os.path.join(orgs_dir, name) for name in os.listdir(orgs_dir) if os.path.isdir(os.path.join(orgs_dir, name))]
    files = []
    for folder in org_folders:
        files += [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.xlsx')]

    logger.info(f'Файлов к загрузке: {len(files)}')
    print(f'🔎 Найдено файлов: {len(files)}')

    frames = []
    for f in files:
        org_name = os.path.basename(os.path.dirname(f))
        logger.info(f'→ Загрузка файла: {os.path.basename(f)} | Организация: {org_name}')
        df = pd.read_excel(f)
        df['organization'] = org_name  # добавляем колонку для агрегации
        frames.append(df)

    if not frames:
        logger.error("❌ Нет файлов для обработки!")
        print("❌ Нет файлов для обработки!")
        return

    df_raw = pd.concat(frames, ignore_index=True)
    df_aggr = aggregate_data(df_raw)
    write_to_excel(df_aggr)
    logger.info('✓ Готово!')
    print('✓ Готово!')

if __name__ == "__main__":
    main()