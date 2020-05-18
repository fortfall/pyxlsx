import logging
import pytest

logger = logging.getLogger(__name__)

@pytest.mark.parametrize("row, amount", [(1, 0), (1, 1), (2, 1), (2, 3), (4, 1)])
def test_insert_rows(ws_with_content, row, amount):
    ws = ws_with_content
    header_row = ws.header_row
    max_row = ws.max_row
    max_col = ws.max_column
    ws.insert_rows(row, amount)
    assert ws.header
    if row <= header_row:
        assert ws.header_row == header_row + amount 
    else:
        assert ws.header_row == header_row
    if row <= max_row:
        assert ws.max_row == max_row + amount
    else:
        assert ws.max_row == row + amount
    assert ws.max_column == max_col

@pytest.mark.parametrize("col, amount", [(1, 0), (1, 1), (2, 1), (6, 3)])
def test_insert_cols(ws_with_content, col, amount):
    ws = ws_with_content
    header_row = ws.header_row
    max_row = ws.max_row
    logger.debug(f"max_row: {max_row}")
    max_col = ws.max_column
    logger.debug(f"max_col: {max_col}")
    ws.insert_cols(col, amount)
    assert ws.header
    assert ws.header_row == header_row
    assert ws.max_row == max_row
    if col <= max_col:
        assert ws.max_column == max_col + amount
    else:
        assert ws.max_column == max_col

@pytest.mark.parametrize("row, amount", [(1, 0), (1, 1), (2, 1), (6, 3)])
def test_delete_rows(ws_with_content, row, amount):
    ws = ws_with_content
    header_row = ws.header_row
    max_row = ws.max_row
    max_col = ws.max_column
    ws.delete_rows(row, amount)
    if header_row >= row and header_row < row + amount:
        assert not ws.header
        assert not ws.header_row
    else:
        assert ws.header
        if header_row < row:
            assert ws.header_row == header_row
        else:
            assert ws.header_row == header_row - amount
    if row > max_row:
        assert ws.max_row == max_row
    elif max_row >= row + amount - 1:
        assert ws.max_row == max_row - amount
    else:
        assert ws.max_row == row - 1
    assert ws.max_column == max_col

@pytest.mark.parametrize("col, amount", [(1, 0), (1, 1), (2, 2), (6, 3)])
def test_delete_cols(ws_with_content, col, amount):
    ws = ws_with_content
    header_row = ws.header_row
    max_row = ws.max_row
    max_col = ws.max_column
    ws.delete_cols(col, amount)
    assert ws.header
    assert ws.header_row == header_row
    if col > max_col:
        assert ws.max_column == max_col
    elif max_col >= col + amount - 1:
        assert ws.max_column == max_col - amount
    else:
        assert ws.max_column == col - 1
    assert ws.max_row == max_row



    