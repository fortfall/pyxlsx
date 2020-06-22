import logging
from openpyxl.cell.read_only import ReadOnlyCell as _ReadOnlyCell, EmptyCell as _EmptyCell
from .Cell import Cell

logger = logging.getLogger(__name__)

class ReadOnlyCell(_ReadOnlyCell):
    __slots__ = ('_cache', '_cache_type')
    def __init__(self, sheet, row, column, value, data_type='n', style_id=0, cache=None, cache_type=None):
        super().__init__(sheet, row, column, value, data_type=data_type, style_id=style_id)
        self._cache = cache
        self._cache_type = cache_type

    @property
    def is_formula(self):
        return Cell.is_formula.__get__(self)
    
    @property
    def cache(self):
        return self._cache
    
    @property
    def cache_type(self):
        return self._cache_type

    @property
    def data(self):
        return Cell.data.__get__(self)
    
    @data.setter
    def data(self, value):
        if self._value is not None:
            raise AttributeError("Cell is read only")
        self._value = value

    @property
    def horizontal(self):
        return Cell.horizontal.__get__(self)
    
    @property
    def vertical(self):
        return Cell.vertical.__get__(self)
    
    @property
    def top(self):
        return Cell.top.__get__(self)
    
    @property
    def bottom(self):
        return Cell.bottom.__get__(self)
    
    @property
    def left(self):
        return Cell.left.__get__(self)
    
    @property
    def right(self):
        return Cell.right.__get__(self)
