import logging
from openpyxl.cell import Cell as _Cell

logger = logging.getLogger(__name__)

class Cell(_Cell):
    @property
    def is_formula(self):
        return self.data_type == 'f' or self.data_type == 'e'
        
    @property
    def formula(self):
        if self.is_formula:
            return self._value
        return None

    @property
    def data(self):
        # out = self._value
        if self.is_formula:
            # result = None
            # if self.parent.data_only_copy is not None:
            #     result = self.parent.data_only_copy.cell(self.row, self.column)._value
            if self.parent.data_only_copy is None \
                    or self.parent.data_only_copy.cell(self.row, self.column).value is None:
                out = self.parent.parent._compute(self.parent.title, self.coordinate)
            else:
                out = self.parent.data_only_copy.cell(self.row, self.column).value
            if out == None:
                out = self._value
        else:
            out = self._value
        if self.parent.header is None or self.row <= self.parent.header_row:
            return out
        if self.parent.use_default:
            if out is None:
                try:
                    out = self.parent.header.get_default(self.column)
                except Exception as e:
                    logger.debug(str(e))
            else:
                default_type = self.parent.header.get_type(self.column)
                if default_type is not None:
                    try:
                        out = default_type(out)
                    except Exception as e:
                        if default_type == int:
                            out = 0
                        elif default_type == float:
                            out = 0.0
                        logger.warning(str(e))
        return out
    
    @data.setter
    def data(self, value):
        self._bind_value(value)
    
    @property
    def horizontal(self):
        return tuple(self.parent.cell(self.row, x) for x in range(self.column, self.parent.max_column + 1))
    
    @horizontal.setter
    def horizontal(self, it):
        for idx, item in enumerate(it):
            self.parent.cell(self.row, self.column + idx, item)
    
    @property
    def vertical(self):
        return tuple(self.parent.cell(x, self.column) for x in range(self.row, self.parent.max_row + 1))
    
    @vertical.setter
    def vertical(self, it):
        for idx, item in enumerate(it):
            self.parent.cell(self.row + idx, self.column, item)
    
    @property
    def top(self):
        if self.row == 1:
            raise IndexError(f"cell({self.row}, {self.column}) has no cell above.")
        return self.parent.cell(self.row - 1, self.column)
    
    @property
    def bottom(self):
        return self.parent.cell(self.row + 1, self.column)
    
    @property
    def left(self):
        if self.column == 1:
            raise IndexError(f"cell({self.row}, {self.column}) has not cell on its left.")
        return self.parent.cell(self.row, self.column - 1)
    
    @property
    def right(self):
        return self.parent.cell(self.row, self.column + 1)
    

        
            
            


            
