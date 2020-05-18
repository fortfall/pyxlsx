import pytest
from pyxlsx import Worksheet

def test_content_column(ws_with_content: Worksheet):
    ws = ws_with_content
    content_column = ws.get_content_column('productName')
    names = set(c.data for c in content_column)
    assert 'pork' in names
    assert 'beef' in names
    assert len(names) == 2