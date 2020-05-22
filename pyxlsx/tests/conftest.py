from collections import OrderedDict
from pathlib import Path
import pytest
from ..__init__ import open_xlsx, new_xlsx, Worksheet, Workbook

@pytest.fixture(params=['pyxlsx/tests/data'])
def data_path(request):
    return Path(request.param)

@pytest.fixture
def new_ws(data_path):
    path = data_path / 'test.xlsx'
    with new_xlsx(path) as wb:
        ws = wb.create_sheet('sheet1')
        yield ws

@pytest.fixture
def wb_read_only(data_path):
    path = data_path / 'read_only.xlsx'
    with open_xlsx(path, read_only=True) as wb:
        yield wb

@pytest.fixture
def ws_with_content(new_ws: Worksheet):
    new_ws.append(
        ["", "", "str('Unknown')", "float(4.5)", "int(500)", "str()"]
    )
    # keys can only be of type str
    content1 = {
        'id': '001',
        'productName': 'pork',
        'productType': 'meat',
        'price': 2.5,
        'weight': 1000,
    }
    content2 = {
        'id': '002',
        'productName': 'beef',
        'productType': 'meat',
        'price': 4.5,
        'weight': 1000,
        'origin': 'Australia'
    }
    new_ws.append_by_header(content1)
    new_ws.append_by_header(content2)
    yield new_ws