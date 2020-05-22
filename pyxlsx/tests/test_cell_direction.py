import pytest
from ..Cell import Cell
from ..open_xlsx import open_xlsx


@pytest.fixture(params=[1, 2])
def row(request):
    return request.param

@pytest.fixture(params=[1, 2])
def column(request):
    return request.param

@pytest.fixture
def cell(new_ws, row, column):
    return new_ws.cell(row, column)

def test_cell_top(cell):
    try:
        top = cell.top
        assert isinstance(top, Cell)
        assert top.column == cell.column
        assert top.row == cell.row - 1
    except Exception as e:
        assert isinstance(e, IndexError)

def test_cell_bottom(cell):
    try:
        bottom = cell.bottom
        assert isinstance(bottom, Cell)
        assert bottom.column == cell.column
        assert bottom.row == cell.row + 1
    except Exception as e:
        assert isinstance(e, IndexError)

def test_cell_left(cell):
    try:
        left = cell.left
        assert isinstance(left, Cell)
        assert left.column == cell.column - 1
        assert left.row == cell.row
    except Exception as e:
        assert isinstance(e, IndexError)

def test_cell_right(cell):
    try:
        right = cell.right
        assert isinstance(right, Cell)
        assert right.column == cell.column + 1
        assert right.row == cell.row
    except Exception as e:
        assert isinstance(e, IndexError)
