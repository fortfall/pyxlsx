import pytest
from pathlib import Path
from openpyxl import load_workbook

@pytest.fixture
def data_path(request):
    return Path(r"/Users/fort/git/python_projects/pyxlsx/pyxlsx/tests/data")


def test_read(data_path):
    path = data_path / 'cache_test_1.xlsx'
    wb = load_workbook(path, data_only=False)
    ws = wb.active
    for row in ws.rows:
        for c in row:
            print(c.value, end=' ')
        print()
