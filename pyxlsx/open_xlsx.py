import os
import logging
from .Workbook import Workbook
from .ExcelReader import ExcelReader

logger = logging.getLogger(__name__)

def open_xlsx(filename, read_only=False) -> Workbook:
    '''
    Open an existing xlsx file.
    Args: 
        filename: path of a xlsx file
        read_only: whether the xlsx file can only be read
    Returns:
        Workbook
    '''
    if not os.path.exists(filename):
        raise FileNotFoundError(f"{filename} not found.")
    reader = ExcelReader(filename, read_only=read_only)
    reader.read()
    wb = reader.wb
    wb.filename = filename
    return wb