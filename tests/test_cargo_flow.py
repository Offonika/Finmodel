import pandas as pd

TAX_DEDUCTIBLE_BY_LOGISTIC = {'Карго': False, 'Белая': True}

# simplified pipeline replicating calculate_cogs_batched logic

def run_pipeline(prod_df, price_df, duty_df, orgs_df, settings):
    price_dict = {str(r['Артикул_поставщика']).strip().upper(): r for _, r in price_df.iterrows()}
    duty_dict = {str(r['Предмет']).strip(): r for _, r in duty_df.iterrows()}
    results = []
    for _, row in prod_df.iterrows():
        org = row['Организация']
        vendor_orig = row['Артикул_поставщика']
        vendor_norm = str(vendor_orig).strip().upper()
        subject = row['Предмет']
        name = row['Название']
        _weight = float(row.get('Вес_брутто', 0) or 0)

        price_row = price_dict.get(vendor_norm, {})
        logistics_mode = price_row.get('Тип_Логистики')
        if not isinstance(logistics_mode, str) or not logistics_mode.strip():
            org_row = orgs_df[orgs_df.iloc[:, 0] == org]
            val = org_row.iloc[0]['Тип_Логистики'] if not org_row.empty else ''
            logistics_mode = 'Белая' if isinstance(val, str) and 'бел' in val.lower() else 'Карго'

        price_val = float(price_row.get('Закуп_Цена', 0) or 0)
        rate = (
            settings['usdRate'] if price_row.get('Валюта') == 'USD' else
            settings['cnyRate'] if price_row.get('Валюта') == 'CNY' else 1
        )
        purchase_rub = price_val * rate

        duty_row = duty_dict.get(subject)
        _duty_rate = 0
        if logistics_mode == 'Белая' and duty_row is not None:
            raw = duty_row.get('Ставка_пошлины') or duty_row.get('Пошлина')
            if raw:
                raw_str = str(raw).replace('%', '').replace(',', '.').strip()
                try:
                    _duty_rate = float(raw_str) if float(raw_str) < 1 else float(raw_str) / 100
                except Exception:
                    _duty_rate = 0


        is_deductible = TAX_DEDUCTIBLE_BY_LOGISTIC.get(logistics_mode, True)
        cogs_mgmt = purchase_rub
        cogs_tax = purchase_rub if is_deductible else 0

        results.append({
            'Организация': org,
            'Артикул_поставщика': vendor_orig,
            'Предмет': subject,
            'Наименование': name,
            'СебестоимостьУпр': round(cogs_mgmt),
            'СебестоимостьНалог': round(cogs_tax),
        })
    return pd.DataFrame(results)

def test_cargo_tax_zero():
    # sample 50 rows with cargo logistics
    prod_df = pd.DataFrame({
        'Организация': ['Org1'] * 50,
        'Артикул_поставщика': [f'ITEM{i}' for i in range(50)],
        'Предмет': ['Cat'] * 50,
        'Название': [f'Name{i}' for i in range(50)],
        'Вес_брутто': [1] * 50,
    })

    price_df = pd.DataFrame({
        'Артикул_поставщика': [f'ITEM{i}' for i in range(50)],
        'Закуп_Цена': [100] * 50,
        'Валюта': ['RUB'] * 50,
        'Тип_Логистики': ['Карго'] * 50,
    })

    duty_df = pd.DataFrame({'Предмет': ['Cat'], 'Ставка_пошлины': [0]})
    orgs_df = pd.DataFrame({'Организация': ['Org1'], 'Тип_Логистики': ['Карго']})
    settings = {
        'cargoRatePerKg': 0,
        'whiteRatePerKg': 0,
        'usdRate': 1,
        'cnyRate': 1,
        'ndsRateWhite': 0.2,
    }

    result = run_pipeline(prod_df, price_df, duty_df, orgs_df, settings)
    assert (result['СебестоимостьНалог'] == 0).all()


def test_management_vs_tax_cogs():
    prod_df = pd.DataFrame({
        'Организация': ['Org1', 'Org1'],
        'Артикул_поставщика': ['IT1', 'IT2'],
        'Предмет': ['Cat', 'Cat'],
        'Название': ['N1', 'N2'],
        'Вес_брутто': [1, 1],
    })

    price_df = pd.DataFrame({
        'Артикул_поставщика': ['IT1', 'IT2'],
        'Закуп_Цена': [100, 100],
        'Валюта': ['RUB', 'RUB'],
        'Тип_Логистики': ['Карго', 'Белая'],
    })

    duty_df = pd.DataFrame({'Предмет': ['Cat'], 'Ставка_пошлины': [0]})
    orgs_df = pd.DataFrame({'Организация': ['Org1'], 'Тип_Логистики': ['Белая']})
    settings = {
        'cargoRatePerKg': 0,
        'whiteRatePerKg': 0,
        'usdRate': 1,
        'cnyRate': 1,
        'ndsRateWhite': 0.2,
    }

    result = run_pipeline(prod_df, price_df, duty_df, orgs_df, settings)
    assert 'СебестоимостьУпр' in result.columns
    assert 'СебестоимостьНалог' in result.columns
    assert result['СебестоимостьУпр'].sum() > result['СебестоимостьНалог'].sum()

