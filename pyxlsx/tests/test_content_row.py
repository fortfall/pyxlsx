import pytest

def test_content_row(ws_with_content):
    ws = ws_with_content
    header = ws.header
    for index, content_row in enumerate(ws.content_rows):
        if content_row['id'] == '001':
            assert content_row['productName'] == 'pork'
            assert content_row[3] == 'meat'
            content_row.update(
                {
                    'id': '003',
                    'desc': 'good pork'
                }
            )
            assert ws.header['desc']
            assert content_row['id'] == '003'
            assert content_row['desc'] == 'good pork'
        if content_row['id'] == '002':
            assert content_row['productName'] == 'beef'
            assert content_row[3] == 'meat'

def test_update_header(ws_with_content):
    ws = ws_with_content
    header = ws.header
    header.update(
        {
            'origin': 'country',
            'desc': None
        }
    )
    for content_row in ws.content_rows:
        if content_row['id'] == '001':
            assert content_row['country'] == None
            content_row.update(
                {
                    'desc': 'good pork'
                }
            )
            assert content_row['desc'] == 'good pork'
        if content_row['id'] == '002':
            assert content_row['country'] == 'Australia'
            assert content_row['desc'] == None