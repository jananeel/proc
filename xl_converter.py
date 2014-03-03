import itertools
import xlrd
import xlwt
import logging
from string import ascii_uppercase

logger.setLevel(logging.INFO)

class XLConverter:
    """
        XLConverter that takes in a inbook, outbook and config mapping file
        on parse() it copies from inbook to outbook, according to the rules in config
    """

    def __init__(self,inbook,outbook,config_file):
        self.inbook = inbook
        self.outbook = outbook
        self.config_file = config_file
        self.conv_map = {}
        self.outsheet_map = {}
        self.parse(config_file)
        self.style = xlwt.XFStyle()
        self.style.alignment.wrap = 1

    def parse(self,config):
        with open(config) as f:
            for line in f.readlines():
                logging.debug(line)
                if not line.strip():
                    continue
                (src,dst) = line.split('=')
                (insheet,irowcol) = src.split(';')
                out_location = dst.split(';')

                if len(out_location) == 1:
                    orowcol = out_location[0]
                    outsheet = insheet
                else:
                    (outsheet,orowcol) = out_location

                input_row_cols = self.generate_row_col(irowcol)
                output_row_cols = self.generate_row_col(orowcol)

                if len(input_row_cols) != len(output_row_cols):
                    logging.error("Error parsing config file: Not matching conversion @ " + line)
                for i,invalue in enumerate(input_row_cols):
                    self.conv_map[(insheet,invalue[0],invalue[1])] = (outsheet,output_row_cols[i][0],output_row_cols[i][1])
                    logging.debug("In (%d,%d) => (%d,%d) " % (invalue[0],invalue[1],output_row_cols[i][0],output_row_cols[i][1]))


    def generate_row_col(self,rowrange):
        """
        Given an excel row col range return all the rows and columns in ascending order of row,col
        given A32:A35
        returns [(31,0),(32,0),(33,0),(34,0)]
        """
        rowranges = rowrange.split(":")
        (row_start,col_start) = self.decode_single_cell(rowranges[0])
        (row_end,col_end) = self.decode_single_cell(rowranges[-1])

        coords = []
        for i in range(row_start,row_end+1):
            for j in range(col_start,col_end+1):
              coords.append((i,j))

        return coords if coords else [(row_start,col_start)]

    def decode_single_cell(self,cell):
        """
        A23 is converted to (22,0)
        """
        col = self.__compute_col(cell)
        row = int("".join([s for s in cell if s.isdigit()])) -1
        return (row,col)

    def __compute_col(self,col):
        col_val = 0
        step = 0
        for alpha in col:
            if alpha not in ascii_uppercase:
                continue
            col_val = col_val + ( step * (len(ascii_uppercase) -1) ) + ascii_uppercase.find(alpha)
            step += 1
        return col_val

    def process(self):
        for key, value in self.conv_map.iteritems():
            insheet = self.inbook.sheet_by_name(key[0])
            (inrow,incol) = (key[1],key[2])
            (outsheet_name,outrow,outcol) = value
            outsheet = self.get_outsheet(outsheet_name)
            # copy the value from insheet to outsheet
            self.copy(insheet,inrow,incol,outsheet,outrow,outcol)

    def get_outsheet(self,outsheet_name):
        if outsheet_name not in self.outsheet_map:
            self.outbook.add_sheet(outsheet_name)
            self.outsheet_map[outsheet_name] = True
        return self.get_sheet_by_name(self.outbook,outsheet_name)

    def get_sheet_by_name(self,book, name):
        """Get a sheet by name from xlwt.Workbook, a strangely missing method.
        Returns None if no sheet with the given name is present.
        """
        # Note, we have to use exceptions for flow control because the
        # xlwt API is broken and gives us no other choice.
        try:
            for idx in itertools.count():
                sheet = book.get_sheet(idx)
                if sheet.name == name:
                    return sheet
        except IndexError:
            return None

    def copy(self,insheet,inrow,incol,outsheet,orow,ocol):
        invalue = insheet.cell_value(rowx=inrow,colx=incol)
        outsheet.write(orow,ocol,invalue,self.style)

    def save(self):
        self.book.save()

    def write_value(self,outsheet_name,rowcol,value):
        (row,col) = self.decode_single_cell(rowcol)
        outsheet = self.get_outsheet(outsheet_name)
        outsheet.write(row,col,value,self.style)
