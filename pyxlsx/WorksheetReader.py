# from openpyxl
import traceback
import logging
from warnings import warn

# compatibility imports
from openpyxl.xml.functions import iterparse

# package imports
from .Cell import Cell  # use reinherited Cell
from openpyxl.cell import MergedCell  # use reinherited MergedCell

from openpyxl.worksheet._reader import WorksheetReader as _WorksheetReader
from .WorksheetParser import WorksheetParser

logger = logging.getLogger(__name__)

class WorksheetReader(_WorksheetReader):
    """
    Create a parser and apply it to a workbook
    """
    def __init__(self, ws, xml_source, shared_strings, data_only, read_only):
        self.ws = ws
        self.parser = WorksheetParser(xml_source, shared_strings, data_only, ws.parent.epoch, ws.parent._date_formats)
        self.tables = []
        self.read_only = read_only

    def bind_cells(self):
        for idx, row in self.parser.parse():
            for cell in row:
                # logger.debug(cell)
                style = self.ws.parent._cell_styles[cell['style_id']]
                c = Cell(self.ws, row=cell['row'], column=cell['column'], style_array=style)
                c._value = cell['value']
                c.data_type = cell['data_type']
                # read cache of formula cell
                if c.data_type == 'f':
                    c._cache = cell['cache']
                    c._cache_type = cell['cache_type']
                self.ws._cells[(cell['row'], cell['column'])] = c
        self.ws.formula_attributes = self.parser.array_formulae
        if self.ws._cells:
            self.ws._current_row = self.ws.max_row # use cells not row dimensions