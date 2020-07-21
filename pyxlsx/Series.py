from typing import Dict, List, Optional, Union, Tuple
import logging
from enum import Enum
# from pyxlsx.Worksheet import Worksheet
from .Cell import Cell

logger = logging.getLogger(__name__)

class InvalidOperationError(Exception):
    pass

class SeriesType(Enum):
    Row = 0
    Column = 1

class Series:
    '''
    Base class for Row and Column
    '''
    def __init__(self, parent, cells, series_type):
        if len(cells) == 0:
            raise ValueError(f"Cannot init Series with empty cells.")
        self._series_type: SeriesType = series_type
        self.parent = parent
        self.cells: Tuple[Cell] = cells
        self.cell_dict: Dict[int, Cell] = self._gen_cell_dict()
    
    def _gen_cell_dict(self):
        if self.series_type == SeriesType.Row:
            return { c.column: c for c in self.cells }
        else:
            return { c.row: c for c in self.cells }
    
    def __len__(self):
        return len(self.cells)
    
    def __iter__(self):
        return iter(self.cells)
    
    def __contains__(self, value):
        for c in self.cells:
            if value == c.data:
                return True
        return False
    
    @property
    def values(self):
        return tuple(c.data for c in self.cells)
    
    @property
    def series_type(self):
        return self._series_type
    
    def __len__(self):
        if self._series_type == SeriesType.Row:
            return self.parent.max_column + 1 - self.parent.min_column
        elif self._series_type == SeriesType.Column:
            return self.parent.max_row + 1 - self.parent.min_row

class ContentRow(Series):
    def __init__(self, parent, cells):
        super().__init__(parent, cells, SeriesType.Row)
        self._min_column: int = self.cells[0].column
        self._max_column: int = self.cells[-1].column
    
    @property
    def row(self):
        return self.cells[0].row

    @property
    def max_column(self):
        return self._max_column
    
    @property
    def min_column(self):
        return self._min_column
    
    def __getitem__(self, key):
        column = self.key_to_column(key)
        try:
            return self.cell_dict[column].data
        except KeyError:
            self.refresh(self.parent.header.max_column - self.max_column)
            return self.cell_dict[column].data
    
    def __setitem__(self, key, value):
        column = self.key_to_column(key)
        self.cell_dict[column].data = value
    
    def refresh(self, increment=0, new_row=None):
        '''
        Refresh cells.
        Args:
            increment: additional columns to include.
            new_row: get cells from different row.
        '''
        if new_row is not None and not isinstance(new_row, int):
            raise TypeError(f"new_row should be of type int or None (got {type(new_row)}.")
        if new_row is not None and new_row <= 0:
            raise ValueError(f"new_row should be positive (got {new_row}).")
        row = new_row or self.row
        self._max_column += increment
        self.cells = tuple(self.parent.cell(row, column) for column in range(self._min_column, self._max_column + 1))
        self.cell_dict = self._gen_cell_dict()

    def append(self, value):
        if value is None:
            return
        self.parent.cell(self.row, self._max_column + 1, value)
        self.refresh(increment=1)
        if not isinstance(self, Header) and self.max_column > self.parent.header.max_column:
            self.parent.header.refresh(self.max_column - self.parent.header.max_column)
    
    def update(self, data_dict, update_header=True):
        for k, v in data_dict.items():
            column = self.parent.header.key_to_column(k)
            if column is not None:
                if column in self.cell_dict:
                    self.cell_dict[column].data = v
                else:
                    self.parent.cell(self.row, column, v)
                    self.refresh(self.parent.header.max_column - self.max_column)
            elif update_header:
                self.parent.header.append(k)
                self.refresh(self.parent.header.max_column - self.max_column)
                self[k] = v
    
    def extend(self, iterable):
        if len(iterable) == 0:
            return
        for incr, value in enumerate(iterable, 1):
            self.parent.cell(self.row, self._max_column + incr, value)
        self.refresh(incr=len(iterable))

    def key_to_column(self, key):
        key_type = type(key)
        if key_type not in {int, str}:
            raise KeyError(f"key type should be int or str (got {type(key)}). Use str to reference by header; use int to reference by index.")
        if key_type == str:
            key = self.parent.header.get_column(key)
        return key
    
    def cell(self, key):
        '''
        Get cell by index with int key or by header with str key.
        return Cell
        '''
        key_type = type(key)
        if key_type not in {int, str}:
            raise KeyError(f"key type should be int or str (got {type(key)}). Use str to reference by header; use int to reference by index.")
        if key_type == int:
            return self.cells[key]
        else:
            column = self.parent.header.get_column(key)
            if column is None:
                raise KeyError(f"key {key} is not in header row")
            try:
                return self.cell_dict[column]
            except KeyError:
                self.refresh(self.parent.header.max_column - self.max_column)
                return self.cell_dict[column]

class Header(ContentRow):
    def __init__(self, parent, cells):
        super().__init__(parent, cells)
        self.column_map: Dict[str, int] = self._init_column_map()
    
    def rebuild(self):
        '''
        Recheck min_row and max_row of Worksheet then reset object with new cell range.
        '''
        self.cells = self.parent[self.parent.header_row]
        self._min_column = self.cells[0].column
        self._max_column = self.cells[-1].column
        self.cell_dict = self._gen_cell_dict()
        self.column_map = self._init_column_map()

    def refresh(self, increment=0, new_row=None):
        '''
        Refresh cells; cell range remains unchanged.
        '''
        super().refresh(increment, new_row)
        self.column_map = self._init_column_map()
    
    def __setitem__(self, key, value):
        super().__setitem__(key, value)
        self.refresh()
    
    def _init_column_map(self):
        cmap = {}
        for c in self.cells:
            cell_value = c.data
            key = str(cell_value) if cell_value is not None else ''
            if key not in cmap:
                cmap[key] = c.column
        return cmap
    
    def get_column(self, key):
        '''
        Get column by header name
        '''
        try:
            return self.column_map[key]
        except:
            return None