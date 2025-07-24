def build_planned_row(data):
    rev = data['rev']
    mp = data['mp']
    fot = data['fot']
    esn = data['esn']
    oth = data['oth']
    ct = data['ct']
    tax = data['tax']
    ebit_tax = rev - (ct + mp + fot + esn + oth)
    return {
        'Себестоимость Налог, ₽': ct,
        'ЧП Налог, ₽': ebit_tax - tax,
    }

def test_planned_row_fields():
    row = build_planned_row({'rev':1000,'mp':100,'fot':50,'esn':10,'oth':40,'ct':300,'tax':20})
    assert row['Себестоимость Налог, ₽'] == 300
    assert row['ЧП Налог, ₽'] == 1000 - (300+100+50+10+40) - 20
