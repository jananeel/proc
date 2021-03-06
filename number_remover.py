import re
import xlwt
from xlutils.copy import copy
import xlrd
import glob
import sys
import logging

logging.basicConfig(level=logging.INFO)

FINDING_SHEET = 0
FINDING_COL = 4
FINDING_DETAIL_COL = 5
def process(inbook, outbook):
    """
    :type book: xlrd.Book book: xlwt.Workbook
    """
    findings_sheet = inbook.sheet_by_name("Findings & CAP")
    for ri in range(2,findings_sheet.nrows):
        finding = unicode(findings_sheet.cell_value(rowx=ri,colx=FINDING_COL)).strip()
        finding_detail = unicode(findings_sheet.cell_value(rowx=ri,colx=FINDING_DETAIL_COL)).strip()

        finding = re.sub(ur'^(\s*(\(?[0-9a-zA-Z]\s*[.):])|([\u2022])\s*)','',unicode(finding),re.UNICODE)
        finding_detail = re.sub(ur'^(\s*\(?[0-9a-zA-Z]\s*[.):]\s*)','',unicode(finding_detail),re.UNICODE)

        outbook.get_sheet(FINDING_SHEET).write(ri,FINDING_COL,finding)
        outbook.get_sheet(FINDING_SHEET).write(ri,FINDING_DETAIL_COL,finding_detail)

    return True


ifiles = glob.glob(sys.argv[1])


for ifile in ifiles:
    logging.info("Removing numbered findings in file:" + ifile)
    try:
        inbook = xlrd.open_workbook(ifile)
        outbook = copy(inbook)
    except xlrd.XLRDError as e:
        logging.exception("Unable to open workbook" + ifile)
        continue

    #todo: exception
    try:
        if not process(inbook,outbook):
            logging.error("Unable to process workbook:" + inbook)
            continue
    except xlrd.XLRDError as e:
        logging.exception("Unable to process file: " + ifile)
        continue

    outbook.save(ifile + ".nonum.xls")
    logging.info('Written file: ' + ifile + ".nonum.xls")

