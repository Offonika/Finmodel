Finmodel – Быстрый старт

НазначениеFinmodel автоматизирует сбор данных с маркетплейсов (Wildberries / Ozon), рассчитывает себестоимость, экономику, налоги и формирует управленческие дашборды в Finmodel.xlsm.

📑 Содержание

Системные требования

Установка

Первый запуск Finmodel.xlsm

Структура листов

Кнопки & скрипты

FAQ

Справка разработчика



1 · Системные требования

ПО

Версия

Windows

10 или новее

Excel

2016 или новее (Win‑32/64)

Python

3.11.x × 64‑bit (авто‑установка)



2 · Установка

2.1 Скачать и распаковать

Finmodel.zip          →  C:\Users\<User>\Documents\Finmodel

2.2 Авто‑установка через setup_finmodel.bat

Двойной клик по setup_finmodel.bat.

В появившемся меню нажмите 4 – «Выполнить все шаги».

Если Python не найден, сценарий скачает и поставит 3.11 × 64‑bit.

🔧  Проблемы с авто‑Python

Если антивирус заблокировал загрузку Python:

Скачайте вручную: https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe

При установке отметьте Add Python to PATH.

После установки запустите setup_finmodel.bat ещё раз и снова выберите 4.

2.3 Что делает скрипт

Шаг

Команда

Описание

1

python -m venv venv

Создаёт виртуальное окружение.

2

pip install -r requirements.txt

Ставит зависимости.

3

xlwings addin install

Добавляет надстройку xlwings в Excel.



3 · Первый запуск Finmodel.xlsm

3.1 Настройка xlwings‑конфигурации

Откройте Finmodel.xlsm.

Вкладка xlwings → Interpreter.

Interpreter – путь к
Finmodel\venv\Scripts\python.exe

PYTHONPATH –
Finmodel\scripts

Отметьте Add workbook to PYTHONPATH.

3.2 Проверка

Нажмите кнопку ℹ️ Version на ленте Управление.Если всплыла «Finmodel ready», настройка верна.



4 · Структура основных листов

Группа

Лист

Назначение

Настройки

Настройки

Ключевые параметры (курсы валют, ставки налогов, периоды).



НастройкиОрганизаций

Токены API, Client‑ID, налоговый режим каждой организации.

Справочники

Номенклатура_WB / Ozon

SKU, размеры, вес, бренд.



ЗакупочныеЦены

Цена, валюта, Тип_логистики (Карго/Белая).



Справочник_льгот

Пошлины, льготы ФОТ и т.д.

Факт

ФинотчетыWB / Ozon

Продажи, возвраты, комиссии.

План

План_Продаж…, План_Выручки…

Автогенерируемые таблицы на год вперёд.

Расчёты

РасчётСебестоимости, РасчётЭкономики…, РасчетПлановыхПоказателей

Итоговые финансовые расчёты.

Отчёты

Дашборд, СводныеПланПродаж

Управленческие сводки и графики.

Полное описание полей – docs/sheets.md (см. репозиторий).



5 · Кнопки и скрипты

5.1 Обновление данных (лента Управление → Загрузка)

#

Кнопка

Скрипт

Итоговый лист

1

ОбновитьЦеныWB

scripts/wb_prices.py

Цены_WB

2

ОбновитьНоменклатуруWB

scripts/import_wb_product_cards.py

Номенклатура_WB

3

ОбновитьФинотчётWB

scripts/wb_report.py

ФинотчетыWB

4

ОбновитьЦеныOzon

scripts/import_ozon_price_info.py

ЦеныOzon

…

…

…

…

5.2 Расчёты

📦 Планы продаж – update_plan_sales.py / update_plan_sales_ozon.py

💵 Планы выручки – update_revenue_plan.py / updateRevenuePlanOzon.py

💰 Себестоимость – calculate_cogs_batched.py

📈 Экономика – update_monthly_scenario_calc.py, economics_table.py

📊 Плановые показатели – fill_planned_indicators.py

🖱️ ЗапуститьВсеРасчёты – макро‑оркестратор, выполняет 4 шага подряд.

⏳ Операции могут занимать до 5–10 мин – ждите всплывающее «✅ Готово».



6 · FAQ

Вопрос

Ответ

Python не найден

Проверьте, что python --version в PowerShell выдаёт 3.11+. Если нет – переустановите Python и перезапустите setup_finmodel.bat.

Ошибка SSL при Wildberries

Убедитесь, что системное время Windows синхронизировано.

Данные не загружаются, токен верный

Проверьте лимиты API (WB → 100 запросов/мин). Либо распределите загрузку – сначала номенклатура, потом отчёты.

Колонки не найдены

Скрипт пишет недостающий список в log/finmodel_*.log. Обычно это означает, что в ручном режиме удалили таблицу заголовков.



7 · Для разработчиков

# запуск скрипта напрямую
cd Finmodel
venv\Scripts\activate
python scripts/calculate_cogs_batched.py --help

Линтинг – ruff check .

Тесты – pytest -q (используется pytest-xlwings для интеграций с Excel)

CI – GitHub Actions: линт + тесты + build‑artifacts (xlsm‑backup).

Шаблон нового скрипта

См. scripts/template_xlwings.py – уже содержит:

get_workbook()

относительные пути (BASE_DIR)

логирование в log/.

© Finmodel 2025  •  Made with 🐍 + 📈 + 💚Вопросы 👉 telegram: @FinmodelSupport

