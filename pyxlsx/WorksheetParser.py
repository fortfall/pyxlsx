import logging
from warnings import warn
from openpyxl.utils import (
    get_column_letter,
    coordinate_to_tuple
)
from openpyxl.cell.text import Text
from openpyxl.utils.datetime import (
    from_excel, from_ISO8601, WINDOWS_EPOCH
)
from openpyxl.worksheet._reader import (
    CELL_TAG, VALUE_TAG, FORMAT_TAG, MERGE_TAG, INLINE_STRING,
    COL_TAG, ROW_TAG, CF_TAG, LEGACY_TAG, PROT_TAG, EXT_TAG, 
    HYPERLINK_TAG, TABLE_TAG, PRINT_TAG, HEADER_TAG, FILTER_TAG,
    VALIDATION_TAG, PROPERTIES_TAG, VIEWS_TAG, FORMULA_TAG, 
    ROW_BREAK_TAG, COL_BREAK_TAG, SCENARIOS_TAG, DATA_TAG, 
    DIMENSION_TAG, CUSTOM_VIEWS_TAG,
    _cast_number
)
from openpyxl.worksheet._reader import WorkSheetParser as _WorksheetParser

logger = logging.getLogger(__name__)

class WorksheetParser(_WorksheetParser):
    def parse_cell(self, element):
        data_type = element.get('t', 'n')
        coordinate = element.get('r')
        self.col_counter += 1
        style_id = element.get('s', 0)
        if style_id:
            style_id = int(style_id)

        if data_type == "inlineStr":
            value = None
        else:
            value = element.findtext(VALUE_TAG, None) or None

        if coordinate:
            row, column = coordinate_to_tuple(coordinate)
        else:
            row, column = self.row_counter, self.col_counter

        # logger.debug(f"after parse: ({row}, {column}) {data_type}: {value}")

        if element.find(FORMULA_TAG) is not None:
            cache_type = data_type
            cache = value
            # logger.debug(f"cache before parse {cache_type}: {cache}")
            data_type = 'f'
            value = self.parse_formula(element)
            cache_type, cache = self._parse_value(element, cache_type, cache, style_id)
            # logger.debug(f"cache before parse {cache_type}: {cache}")
            # logger.debug(f"formula after parse: ({row}, {column}) {data_type}: {value}")
            return {
                'row':row, 
                'column':column, 
                'value':value, 
                'data_type':data_type, 
                'style_id':style_id,
                'cache_type': cache_type,
                'cache': cache
            }

        else:
            data_type, value = self._parse_value(element, data_type, value, style_id)
            # logger.debug(f"after parse: ({row}, {column}) {data_type}: {value}")
            return {'row':row, 'column':column, 'value':value, 'data_type':data_type, 'style_id':style_id}

    def _parse_value(self, element, data_type, value, style_id):
        if value is not None:
            if data_type == 'n':
                value = _cast_number(value)
                if style_id in self.date_formats:
                    data_type = 'd'
                    try:
                        value = from_excel(value, self.epoch)
                    except ValueError:
                        msg = """Cell {0} is marked as a date but the serial value {1} is outside the limits for dates. The cell will be treated as an error.""".format(coordinate, value)
                        warn(msg)
                        data_type = "e"
                        value = "#VALUE!"
            elif data_type == 's':
                value = self.shared_strings[int(value)]
            elif data_type == 'b':
                value = bool(int(value))
            elif data_type == "str":
                data_type = "s"
            elif data_type == 'd':
                value = from_ISO8601(value)

        elif data_type == 'inlineStr':
                child = element.find(INLINE_STRING)
                if child is not None:
                    data_type = 's'
                    richtext = Text.from_tree(child)
                    value = richtext.content

        return (data_type, value)
        