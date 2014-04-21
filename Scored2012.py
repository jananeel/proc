import sys
import xlwt
import xlrd
from batch_file_reader import BatchFileProcessor
from audit_mapper_2012 import AuditScoreMapper2012
from xl_converter import XLConverter

AUDIT_SHEET="Sheet1"
AUDIT_GENERAL_INFO="Audit General Info"
SCORES_SUFFIX = "scores.xls"
NO_CONFIG = "./empty_config"
SCORES = "Scores"
FACILITY_PROFILE = "Facility Profile"

class Scores2012(BatchFileProcessor):
    def __init__(self,pattern,audit_file):
        super(Scores2012,self).__init__(pattern,SCORES_SUFFIX)
        self.dec = XLConverter(None,None,NO_CONFIG)
        self.audit_mapper = AuditScoreMapper2012(audit_file,AUDIT_SHEET)
        self.audit_mapper.process()

    def transform(self,inbook,outbook):
        (r,c) = self.dec.decode_single_cell('B1')
        facility_id = inbook.sheet_by_name(FACILITY_PROFILE).cell_value(rowx=r,colx=c)
        (r,c) = self.dec.decode_single_cell('B5')
        audit_date = inbook.sheet_by_name(AUDIT_GENERAL_INFO).cell_value(rowx=r,colx=c)

        out_sheet = outbook.add_sheet(SCORES)
        preferred_list = self.audit_mapper.get_preferred_cols(facility_id,audit_date)

        if( preferred_list):
            for r,row in enumerate(preferred_list):
                for c,col in enumerate(row):
                    out_sheet.write(r,c,label=col)

        return True

scores = Scores2012(sys.argv[1],sys.argv[2])
scores.process()

