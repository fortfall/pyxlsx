import pytest

def test_is_formula(wb_read_only):
    ws = wb_read_only['is_formula']
    for row in ws.rows:
        cell1 = row[0]
        cell2 = row[1]
        if cell1.data != None:
            assert cell2.is_formula
        else:
            assert not cell2.is_formula
