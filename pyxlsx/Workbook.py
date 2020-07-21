from typing import Union, List
from pycel.excelcompiler import ExcelCompiler
from openpyxl import Workbook as _Workbook
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.chartsheet import Chartsheet
from openpyxl.workbook.defined_name import DefinedNameList
from openpyxl.packaging.core import DocumentProperties
from openpyxl.packaging.relationship import RelationshipList
from openpyxl.workbook.protection import DocumentSecurity
from openpyxl.workbook.properties import CalcProperties
from openpyxl.workbook.views import BookView
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900
from .Worksheet import Worksheet


class Workbook(_Workbook):
    filename: str = None
    def __init__(self, filename=None, read_only=False, iso_dates=False):
        self.filename = filename
        self._read_only = read_only
        self._formula_calculator = None
        self._sheets = []
        self._pivots = []
        self._active_sheet_index = 0
        self.defined_names = DefinedNameList()
        self._external_links = []
        self.properties = DocumentProperties()
        self.security = DocumentSecurity()
        self.__write_only = False
        self.shared_strings = IndexedList()

        self._setup_styles()

        self.loaded_theme = None
        self.vba_archive = None
        self.is_template = False
        self.code_name = None
        self.epoch = CALENDAR_WINDOWS_1900
        self.encoding = "utf-8"
        self.iso_dates = iso_dates

        if not self.write_only:
            self._sheets.append(Worksheet(self))

        self.rels = RelationshipList()
        self.calculation = CalcProperties()
        self.views = [BookView()]
    def __getitem__(self, key) -> Union[Worksheet, Chartsheet]:
        return super().__getitem__(key)
    
    def create_sheet(self, title=None, index=None) -> Worksheet:
        new_ws = Worksheet(parent=self, title=title)
        self._add_sheet(sheet=new_ws, index=index)
        return new_ws
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_value, traceback):
        if self._read_only or self.filename is None:
            self.close()
        else:
            self.save(self.filename)
    
    def save(self, filename=None):
        if self._read_only:
            raise Exception("Workbook can't be saved (read only mode).")
        if filename is not None:
            super().save(filename)
        elif self.filename is not None:
            super().save(self.filename)
    
    def _compute(self, sheet: str, address: str):
        if self._formula_calculator is None:
            self._init_calculator()
        return self._formula_calculator.evaluate(f"{sheet}!{address}")
    
    def _init_calculator(self):
        self._formula_calculator = ExcelCompiler(excel=self)

    @property
    def worksheets(self):
        """A list of sheets in this workbook

        :type: list of :class:`openpyxl.worksheet.worksheet.Worksheet`
        """
        return [s for s in self._sheets if isinstance(s, (Worksheet, ReadOnlyWorksheet, WriteOnlyWorksheet))]
