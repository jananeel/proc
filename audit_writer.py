import xlwt
class AuditReportWriter:
    ROW = 0
    ID = 1
    FINDING = 2
    FINDING_DETAIL = 3
    COMMENT = 4
    TEXT_STYLE = xlwt.XFStyle()
    TEXT_STYLE.alignment.wrap = 1
    TEXT_STYLE_ERROR = xlwt.XFStyle()
    TEXT_STYLE_ERROR.alignment.wrap = 1
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = xlwt.Style.colour_map['lime']
    TEXT_STYLE_ERROR.pattern = pattern

    HEADER_STYLE = xlwt.XFStyle()
    HEADER_STYLE.alignment.wrap = 1

    def __init__(self,xlname,sname):
        self.xlname = xlname
        self.sheet_name = sname
        self.workbook = xlwt.Workbook()
        self.worksheet = self.workbook.add_sheet(self.sheet_name)
        self.write_rw = 0
        self.write_row("Row #","ID", "Finding","Finding Detail", "Comment")
        self.write_rw +=1

        #set col width
        self.worksheet.col(self.ROW).width = 256 * 3 #3 column width
        self.worksheet.col(self.ID).width = 256 * 8
        self.worksheet.col(self.FINDING).width = 256 * 50
        self.worksheet.col(self.FINDING_DETAIL).width = 256 * 60
        self.worksheet.col(self.COMMENT).width = 256 * 10
        
    def save(self):
        self.workbook.save(self.xlname)

    def write_row(self,row_num,id,finding,detail,comment):
        style = self.TEXT_STYLE
        if comment.startswith('Cond4'):
            style = self.TEXT_STYLE_ERROR
        self.worksheet.write(self.write_rw,self.ROW,row_num,style)
        self.worksheet.write(self.write_rw,self.ID,id,style)
        self.worksheet.write(self.write_rw,self.FINDING,finding,style,)
        self.worksheet.write(self.write_rw,self.FINDING_DETAIL,detail,style)
        self.worksheet.write(self.write_rw,self.COMMENT,comment,style)
        self.write_rw +=1
