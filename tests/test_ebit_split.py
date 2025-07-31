from scripts.fill_planned_indicators import _calc_row


def test_mp_excluded_from_mgmt_but_tax_with_gross():
    row = _calc_row(
        revN=1000,
        mpNet=200,
        cost_sales=300,
        cost_tax=360,
        fot=0,
        esn=0,
        oth=0,
        mode='Доходы-Расходы',
        mpGross=240,
    )
    assert row['EBITDA, ₽'] == 500
    assert row['Расчет_базы_налога'] == 400
