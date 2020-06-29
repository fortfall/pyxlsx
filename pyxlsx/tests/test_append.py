from ..__init__ import new_xlsx, open_xlsx

def test_append(data_path):
    path = data_path / 'test_append.xlsx'
    langs = {
        'EN': 'English',
        'CN': 'Chinese'
    }
    words = {
        'EN': 'Hello',
        'CN': '你好'
    }
    with new_xlsx(path) as wb:
        ws = wb.active
        ws.append_by_header(langs)
        ws.append_by_header(words)
    
    with open_xlsx(path) as wb:
        ws = wb.active
        assert ws.cell(1, 1).data == 'EN'
        assert ws.cell(1, 2).data == 'CN'
        assert ws.cell(2, 1).data == 'English'
        assert ws.cell(2, 2).data == 'Chinese'
        assert ws.cell(3, 1).data == 'Hello'
        assert ws.cell(3, 2).data == '你好'

