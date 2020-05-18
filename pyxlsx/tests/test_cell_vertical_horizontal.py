import pytest


@pytest.fixture(params=[(1, 1)])
def write_vertical_cells(new_ws, request):
    row, column = request.param
    cell = new_ws.cell(row, column)
    for x in range(10):
        new_ws.cell(row + x, column, x)
    yield cell

@pytest.fixture(params=[(1, 1)])
def write_horizontal_cells(new_ws, request):
    row, column = request.param
    cell = new_ws.cell(row, column)
    for x in range(10):
        new_ws.cell(row, column + x, x)
    yield cell

@pytest.fixture(params=[(1, 1)])
def cell(new_ws, request):
    row, column = request.param
    return new_ws.cell(row, column)

def test_vertical_read(write_vertical_cells):
    for index, c in enumerate(write_vertical_cells.vertical):
        assert c.data == index

def test_vertical_write(cell):
    cell.vertical = (x + 10 for x in range(15))
    for index, c in enumerate(cell.vertical):
        assert c.data == index + 10

def test_horizontal_read(write_horizontal_cells):
    for index, c in enumerate(write_horizontal_cells.horizontal):
        assert c.data == index

def test_horizontal_write(cell):
    cell.horizontal = range(15)
    for index, c in enumerate(cell.horizontal):
        assert c.data == index