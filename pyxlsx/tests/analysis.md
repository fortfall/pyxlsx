* Workbook
  * add __enter__
  * add __exit__
  * overload save
  * add filename
  * add formula_calculator
* Worksheet
  * add use_default, _use_default
  * header, _header
  * read_copy
  * row_offset, _row_offset
  * content_rows
  * raw_content_rows
  * content_cols
  * trim
  * trim_rows
  * trim_cols
  * overload delete_rows
  * overload delete_cols
  * overload insert_rows
  * overload insert_cols

* ReadOnlyWorksheet
* Cell
* ReadOnlyCell
* ExcelReader
* WorkbookParser
* WorksheetReader
* WorksheetParser

```python
def load_workbook():
    reader = ExcelReader()
    reader.read()
    return reader.wb

class ExcelReader:
    def read_manifest():
        self.package = Manifest.from_tree(root)

    def read_workbook():
        self.parser = WorkbookParser()
        self.parser.parse()
        self.wb = self.parser.wb
        self.wb._read_only = self.read_only
        self._data_only = self.data_only

    def read_worksheet():
        for sheet in self.parser.find_sheets():
            if self.read_only:
                ws = ReadOnlyWorksheet() # from archive, using WorksheetParser
                # ReadOnlyCell EmptyCell instantiated
                self.wb._sheets.append(ws)
            else:
                # Worksheet or WriteOnlyWorksheet
                ws = self.wb.create_sheet(sheet.name) 
                ws_parser = WorksheetReader(archive_file_handler)
                ws_parser.bind_all()  # bind ws with content from archive
                # Cell instantiated

class WorkbookParser:
    def __init__():
        self.wb = Workbook()

    def parse():
        package = WorkbookPackage.from_tree(node)
        self.sheets = package.sheets


```
