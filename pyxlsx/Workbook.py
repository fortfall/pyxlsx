from typing import Union, List
from pycel.excelcompiler import ExcelCompiler
from openpyxl import Workbook as _Workbook
from openpyxl.chartsheet import Chartsheet
from .Worksheet import Worksheet

class Workbook(_Workbook):
    filename: str = None
    def __init__(self, filename=None, read_only=False, use_default=False):
        super().__init__(write_only=False, iso_dates=False)
        self.filename = filename
        self._read_only = read_only
        self._use_default = use_default
        self._formula_calculator = None
    
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