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
    # data_only copy is opened for reading cached formula results
    # reader_data_only = ExcelReader(filename, read_only=read_only, data_only=True)
    # reader_data_only.read()
    # wb_data_only = reader_data_only.wb
    # for name in wb.sheetnames:
    #     wb[name].data_only_copy = wb_data_only[name]
    return wb