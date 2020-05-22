import logging
from ..__init__ import Worksheet

logger = logging.getLogger(__name__)

def test_formula_calc(new_ws: Worksheet):
    ws = new_ws
    ws['A1'] = 1
    ws['A2'] = 2
    ws['a3'] = 3
    ws['a4'] = 4

    ws['B1'] = 2
    ws['B2'] = 4
    ws['B3'] = 6
    ws['B4'] = 8
    for row in range(1, 5):
        ws.cell(row, 3).data = "=A{0}+B{0}".format(row)  
        ws.cell(row, 4).data = "=vlookup(C{0}, A1:B4, 2, FALSE)".format(row)  

    assert ws['c1'].is_formula
    assert ws['c1'].cache is None
    assert ws['c1'].cache_type is None

    assert ws['c2'].is_formula
    assert ws['c2'].cache is None
    assert ws['c2'].cache_type is None

    assert ws['c3'].is_formula
    assert ws['c3'].cache is None
    assert ws['c3'].cache_type is None

    assert ws['c4'].is_formula
    assert ws['c4'].cache is None
    assert ws['c4'].cache_type is None

    assert ws['d1'].is_formula
    assert ws['d1'].cache is None
    assert ws['d1'].cache_type is None

    assert ws['d2'].is_formula
    assert ws['d2'].cache is None
    assert ws['d2'].cache_type is None

    assert ws['d3'].is_formula
    assert ws['d3'].cache is None
    assert ws['d3'].cache_type is None

    assert ws['d4'].is_formula
    assert ws['d4'].cache is None
    assert ws['d4'].cache_type is None

    # logger.debug(ws['C1'].data)
    assert ws['C1'].data == 3
    assert ws['C2'].data == 6
    assert ws['C3'].data == 9
    assert ws['C4'].data == 12

    assert ws['D1'].data == 6
    assert ws['D2'].data == '#N/A'
    assert ws['D3'].data == '#N/A'
    assert ws['D4'].data == '#N/A'
