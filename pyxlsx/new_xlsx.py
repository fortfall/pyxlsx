from .Workbook import Workbook

def new_xlsx(filename=None) -> Workbook:
    '''
    Create a new xlsx file.
    Args:
        filename: path to save the new xlsx file.
    Returns:
        Workbook
    '''
    wb = Workbook(filename)
    return wb