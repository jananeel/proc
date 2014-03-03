import sys
import utils
import xlrd
import glob
import logging
from xl_converter import XLConverter
from audit_writer import AuditReportWriter
from audit_parser import AuditParser

logging.basicConfig(level=logging.DEBUG)

FINDING_TYPE = "Finding"  #value for the type field, in case of findings
OBSERV_TYPE = "Observ" #value for the type field, in case of observations

ifiles = glob.glob(sys.argv[1])

for ifile in ifiles:
    logging.info("Processing file:" + ifile)
#print "The number of worksheet is: %d" % book.nsheets
    try:
        book = xlrd.open_workbook(ifile,encoding_override="cp1252")
        sheet = book.sheet_by_name("Findings & CAP")
    except xlrd.XLRDError as e:
        logging.exception("Unable to open -Findings & CAP- in file: " + ifile )
        continue

    ar = AuditReportWriter(ifile+'.xls','Findings & CAP')

    for rx in range(2,sheet.nrows):
        finding = sheet.cell_value(rowx=rx,colx=4) #Column E, row = rx
        #print finding;
        if finding.lower() == "n/a" or finding.strip() == '':
            continue

        findingdetail1 = sheet.cell_value(rowx=rx,colx=5) #Column F, row=rx
        findingdetail2 = sheet.cell_value(rowx=rx,colx=6) #Column G, row=rx
        id = sheet.cell_value(rowx=rx,colx=1) #Column B, row=rx
        cat = sheet.cell_value(rowx=rx,colx=2), #Column C, row=rx
        cause = sheet.cell_value(rowx=rx,colx=8), #Column I

        try:
            findings = finding.strip().splitlines();  #strip all extra whitespaces and new lines

            d1 = findingdetail1.strip()  #remove leading whitespace
            d2 = findingdetail2.strip()  #remove leading whitespace

            if d1.lower() == 'n/a' or d1.strip() == '' : d1 = ''
            if d2.lower() == 'n/a'or d2.strip() == '' : d2 = ''
        except:
            logging.exception("Error reading, row: " + str(rx))
            #try to write row back
            ar.write_row(rx+1,id,"ERROR","ERROR",finding,unicode(findingdetail1)+"\n" + unicode(findingdetail2),cause,"Cond4: Neelansha Madam. Please look" + str(sys.exc_info()))
            logging.exception(unicode(findingdetail1)+"\n" + unicode(findingdetail2))
            continue


        (fhdr,numbered_findings,fob) = AuditParser("\n".join(findings)).parse()
        (fdhdr,numbered_finding_details,fdob) = AuditParser(d1+'\n'+d2).parse(fob)

        exd_observations = []
        exd_observations.append(fdhdr)
        exd_observations.extend(numbered_finding_details)

        #Check for numeric matches - code matched as Cond3
        if( numbered_findings and numbered_finding_details and len(numbered_findings) == len(numbered_finding_details)):
            for i in range(len(numbered_findings)):
                ar.write_row(rx+1,id,cat,FINDING_TYPE,fhdr + numbered_findings[i] ,fdhdr + numbered_finding_details[i],cause,"Cond3: Numeric match")
        elif numbered_findings and (not numbered_finding_details or len(numbered_findings) != len(numbered_finding_details)):
            for find in numbered_findings:
                ar.write_row(rx+1,id,cat,FINDING_TYPE,fhdr + find,"".join(exd_observations),cause,"Cond2.1")
        elif not numbered_findings:
            ar.write_row(rx+1,id,cat,FINDING_TYPE,fhdr,"".join(exd_observations),cause,"Cond1")
        else:
            #Cond4 - no such match. Dump findings and finding details and notify
            ar.write_row(rx+1,id,cat,FINDING_TYPE,finding,(d1 + "\n" + d2),cause,"Cond4: Neelansha Please look at it!")

        if fob and not fdob:
            ar.write_row(rx+1,id,cat,OBSERV_TYPE,fob,d1 + "\n" + d2,cause,"Cond5.1 observ")
        elif fob and fdob:
            ar.write_row(rx+1,id,cat,OBSERV_TYPE,fob,fdob,cause,"Cond5 observ")

    # #map facility profile
    # try:
    #     facility_profile = book.sheet_by_name("Facility Profile")
    # except xlrd.XLRDError as e:
    #     logging.exception("Unable to open -FacilityProfile- in file: " + ifile)
    #     continue

    # writer = XLWriter(ar.get_workbook(),"Facility Profile")

    # #copy A3:A19 to A2:A18 and B3:B19 to B2:B18
    # tgt_row = 1
    # tgt_col = 1
    # for rx in range(2,19):
    #     writer.copy(facility_profile,rx,0,tgt_row,tgt_col)
    #     writer.copy(facility_profile,rx,1,tgt_row,tgt_col+1)
    #     tgt_row += 1

    # tgt_row = 18
    # tgt_col = 1
    # for rx in range(22,30):
    #     writer.copy(facility_profile,rx,0,tgt_row,tgt_col)
    #     tgt_row += 1

    # tgt_row = 26
    # tgt_col = 0
    # for rx in range(31,38):
    #     writer.copy(facility_profile,rx,0,tgt_row,tgt_col)
    #     tgt_row += 1

    # writer.copy(facility_profile,39,0,33,0)

    # tgt_row = 18
    # for rx in range(22,30):
    #     tgt_col = 2
    #     for cx in range(1,6):
    #         writer.copy(facility_profile,rx,cx,tgt_row,tgt_col)
    #         tgt_col +=1
    #     tgt_row +=1

    # tgt_row = 26
    # for rx in range(31,37):
    #     tgt_col = 2
    #     for cx in range(1,6):
    #         writer.copy(facility_profile,rx,cx,tgt_row,tgt_col)
    #         tgt_col +=1
    #     tgt_row +=1

    # writer.copy(facility_profile,37,1,32,1)
    # writer.copy(facility_profile,39,1,33,1)

    # END Facility Profile

    converter = XLConverter(book,ar.get_workbook(),"./config")
    converter.process()

    ar.save()
    logging.info('Written file: ' + ifile + ".xls")
