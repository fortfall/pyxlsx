### Introduction
A package to read/write xlsx worksheet like dictionary, based on openpyxl.

### Installation
* from pip
```python
pip install pyxlsx
```
* download package and run setup.py
```python
python setup.py install
```

### Usage
* Create a new xlsx file and write to it
```python
from pyxlsx import new_xlsx

with new_xlsx(filename) as wb:
    ws = wb.create_sheet('sheet1')
    # some operations

# or
wb = new_xlsx()
ws = wb.creat_sheet('sheet1')  # create a new sheet with name 'sheet1'
# some operations
wb.save(filename)
```
* Open an existing xlsx file
```python
from pyxlsx import open_xlsx

with open_xlsx(filename) as wb:
    ws1 = wb.active  # get active sheet
    ws2 = wb['sheet2']
    # some operations

# or
wb = open_xlsx(filename)
ws = wb['sheet2']
# some operations
wb.save()
# to save as another file
wb.save(another_filename)
```
* Append rows to a worksheet
```python
ws = wb['sheet1']
ws.append(
    ["", "", "str('Unknown')", "float(4.5)", "int(500)", "str()"]
)
# keys can only be of type str
content1 = {
    'id': '001',
    'productName': 'pork',
    'productType': 'meat',
    'price': 2.5,
    'weight': 1000,
}
content2 = {
    'id': '002',
    'productName': 'beef',
    'productType': 'meat',
    'price': 4.5,
    'weight': 1000,
    'origin': 'Australia'
}
# header is auto-generated from keys of the dict the first time append_by_header is called.
ws.append_by_header(content1)  
# new header name will be append to header if append_header is True (default value)
ws.append_by_header(content2)  
# below is the result of writing operation
```
||A|B|C|D|E|F|
|:---:|:---:|:---:|:---:|:---:|:---:|:---|
|1|||str('Unknown')|float(30)|int(0)|str()
|2|id|productName|productType|price|weight|origin
|3|001|pork|meat|2.5|1000|
|4|002|beef|meat|4.5|1000|Australia
* Read from / write to a worksheet by row
  Note: if there are duplicate header names, only the first would be used.
```python
ws = wb['sheet1']
assert ws.header is None
ws.header_row = 2  # set the second row as worksheet header row
assert ws.header is not None

for row in ws.content_rows:  # starting from row just below header row
    print(row[1])  # row cell value can accessed by column number, if key is of type int
    print(row['productName'])  # row cell value can be accessed by header name, if key is of type of str
    print(row['price'])  
    if row['productName'] == 'pork':
        row[1] = '003'  # change pork id to '003'
        row['price'] = 3.5  # change pork price to 3.5
# output as below
# '001'
# 'pork'
# 2.5
# '002'
# 'beef'
# 4.5
```
* Read from a worksheet by column
```python
ws = wb['sheet1']
ws.header_row = 2
# get a full column
column_cells = ws['B']
for c in column_cells:
    print(c.data)  # 'pork', 'beef'

# get a content column (containing only cells below header) by header name, 
# if key is of type str
name_column = ws.get_content_column('productName')
for c in name_column:
    print(c.data)  # 'pork', 'beef'

# get a content column by column number,
# if key is of type int
name_column = ws.get_content_column(2)
for c in name_column:
    print(c.data)  # 'pork', 'beef'
```
* Read cell directly from Worksheet, Header, ContentRow
```python
ws = wb['sheet1']
ws.header_row = 2
# access a cell by coordinate (row, column)
cell = ws.cell(2, 2)
print(cell.data)  # 'productName'

# access a cell by header name if key is of type str
cell = ws.header.cell('productName')
print(cell.data)  # 'productName'
# access a cell by column number
cell = ws.header.cell(1)
print(cell.data)  # 'id'

for row in ws.content_rows:
    cell = row.cell(1)  # '001', '002'
    print(cell.data)
    cell = row.cell('productName')
    print(cell.data)  # 'pork', 'beef'
```

* Read adjacent cells of a certain cell
```python
cell = ws.cell(2, 2)
print(cell.top.data)  # "str('Unknown')"
print(cell.left.data)  # 'id'
print(cell.right.data)  # 'productType'
print(cell.bottom.data)  # 'pork'

for c in cell.vertical:
    print(c.data)  # 'productName', 'pork', 'beef'

for c in cell.horizontal:
    print(c.data)  # 'productName', 'productType', 'price', 'weigth', 'origin'
```