import xlrd

FACILITY_ID_COL=0
FACILITY_NAME_COL=10
YEAR_COL = 7
AUDIT_MAP_SHEET = 'all audits'

class FacilityMapper:
    """
    This class provides the facility id to name mapping
    """

    def __init__(self,config_file):
        """
        Takes the input file that contains the facility mapping and stores the map of facility-name to facility id
        """
        self.book = xlrd.open_workbook(config_file)
        self.sheet = self.book.sheet_by_name(AUDIT_MAP_SHEET)
        self.facility_map={}
        self.create_map()

    def create_map(self):
        for rx in range(self.sheet.nrows):
            facility_id = self.sheet.cell_value(rowx=rx,colx=FACILITY_ID_COL)
            year = self.sheet.cell_value(rowx=rx,colx=YEAR_COL)
            facility_name = self.sheet.cell_value(rowx=rx,colx=FACILITY_NAME_COL).strip()
            #if (facility_name,year) in self.facility_map:
            #    raise Exception("Input exception: Duplicate value: %s,%d " %(facility_name, year))
            self.facility_map[(facility_name,year)]=facility_id


    def get_facility_id(self,facility_name,year):
        """
        Given a facility name, gives the corresponding facility id
        """
        if (facility_name.strip(),year) in self.facility_map:
            return self.facility_map[(facility_name.strip(),year)]
        return None
