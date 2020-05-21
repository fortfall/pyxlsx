import logging
from openpyxl.cell.read_only import ReadOnlyCell as _ReadOnlyCell, EmptyCell as _EmptyCell
from .Cell import Cell

logger = logging.getLogger(__name__)

class ReadOnlyCell(_ReadOnlyCell):
    @property
    def is_formula(self):
        return Cell.is_formula.__get__(self)
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