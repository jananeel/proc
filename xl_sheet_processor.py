import xlrd
from xl_converter import XLConverter

class XLSheetProcessor(object):
    """
    This class abstracts implements the base functionality needed for processing an excel sheet, row by row
    """

    def __init__(self,file,sheet):
        """
        Takes the input file and the sheet name
        """
        self.book = xlrd.open_workbook(file)
        self.sheet = self.book.sheet_by_name(sheet)
        self.dec = XLConverter(None,None,"./empty_config")

    def process(self):
        for r in range(self.sheet.nrows):
            self.process_row(self.sheet.row_values(r))
        
    def process_row(self,row_values):
        pass
    
    def read_col(self,row_values,col):
        c = self.dec.decode_col(col)
        return row_values[c]

