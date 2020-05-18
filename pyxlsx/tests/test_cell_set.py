import pytest

def test_cell_set(new_ws):
    for x in range(1, 10):
        new_ws.cell(1, x).data = x
    for x in range(1, 10):
        cell = new_ws.cell(1, x)
        assert cell.data == x
