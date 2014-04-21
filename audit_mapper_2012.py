import collections
from xl_sheet_processor import XLSheetProcessor

class AuditScoreMapper2012(XLSheetProcessor):
    "This class abstracts out the Audit Score mappers"

    def __init__(self,file,sheet):
        super(AuditScoreMapper2012,self).__init__(file,sheet)
        self.audit_map = collections.defaultdict(list)

    def process_row(self,row_values):
        fac_id = self.read_col(row_values,'H')
        audit_date = self.read_col(row_values,'L')

        a_col = self.read_col(row_values,'A')
        b_col = self.read_col(row_values,'B')
        c_col = self.read_col(row_values,'C')
        d_col = self.read_col(row_values,'D')
        p_col = self.read_col(row_values,'P')

        self.audit_map[(fac_id,audit_date)].append([a_col,b_col,c_col,d_col,p_col])

    def get_preferred_cols(self,facility_id,audit_date):
        return self.audit_map[(facility_id,audit_date)]

