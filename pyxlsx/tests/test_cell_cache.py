import logging
from pathlib import Path
from ..open_xlsx import open_xlsx
from ..new_xlsx import new_xlsx

logger = logging.getLogger(__name__)

def test_cell_cache(data_path):
    path = data_path / 'cache_test.xlsx'
    for read_only in [True, False]:
        logger.debug(f"read_only: {read_only}")
        wb = open_xlsx(path, read_only=read_only)
        ws: Worksheet = wb.active
        c1 = ws['c1']
        c2 = ws['c2']
        c3 = ws['c3']
        c4 = ws['c4']
        c5 = ws['c5']
        c6 = ws['c6']

        d1 = ws['d1']
        d2 = ws['d2']
        d3 = ws['d3']
        d4 = ws['d4']
        d5 = ws['d5']
        d6 = ws['d6']

        assert c1.is_formula
        assert c1.cache is not None
        assert c1.cache_type == 'n'

        assert c2.is_formula
        assert c2.cache is not None
        assert c2.cache_type == 'n'

        assert c3.is_formula
        assert c3.cache is not None
        assert c3.cache_type == 'n'

        assert c4.is_formula
        assert c4.cache is not None
        assert c4.cache_type == 'n'

        assert not c5.is_formula
        assert c5.cache is None
        assert c5.cache_type is None

        assert not c6.is_formula
        assert c6.cache is None
        assert c6.cache_type is None

        assert d1.is_formula
        assert d1.cache is not None
        assert d1.cache_type == 'n'

        assert d2.is_formula
        assert d2.cache is not None
        assert d2.cache_type == 'e'

        assert d3.is_formula
        assert d3.cache is not None
        assert d3.cache_type == 'e'

        assert d4.is_formula
        assert d4.cache is not None
        assert d4.cache_type == 'e'

        assert d5.is_formula
        assert d5.cache is not None
        assert d5.cache_type == 's'

        assert d6.is_formula
        assert d6.cache is not None
        assert d6.cache_type == 'e'
