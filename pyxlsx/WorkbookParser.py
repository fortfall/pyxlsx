from openpyxl.reader.workbook import WorkbookParser as _WorkbookParser
from .Workbook import Workbook

class WorkbookParser(_WorkbookParser):
    def __init__(self, archive, workbook_part_name, keep_links=True):
        self.archive = archive
        self.workbook_part_name = workbook_part_name
        self.wb = Workbook()
        self.keep_links = keep_links
        self.sheets = []