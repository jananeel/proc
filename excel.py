import sys
import utils
import xlrd
import glob
import logging
from xl_converter import XLConverter
from audit_writer import AuditReportWriter
from audit_parser import AuditParser
from facility_profile_mapper import FacilityMapper

logging.basicConfig(level=logging.INFO)

FINDING_TYPE = "Finding"  #value for the type field, in case of findings
OBSERV_TYPE = "Observ" #value for the type field, in case of observations

ifiles = glob.glob(sys.argv[1])
facility_profile_file = sys.argv[2]
config_file = sys.argv[3]

#read facility profile config
facility_mapper  = FacilityMapper(facility_profile_file)


for ifile in ifiles:
    logging.info("Processing file:" + ifile)
#print "The number of worksheet is: %d" % book.nsheets
    try:
        book = xlrd.open_workbook(ifile,encoding_override="cp1252")
        sheet = book.sheet_by_name("Findings & CAP")
        facility_profile = book.sheet_by_name("Facility Profile")
    except xlrd.XLRDError as e:
        logging.exception("Unable to open -Findings & CAP- in file: " + ifile )
        continue

    ar = AuditReportWriter(ifile+'.xls','Findings & CAP')

    for rx in range(2,sheet.nrows):
        try:

            finding = sheet.cell_value(rowx=rx,colx=4) #Column E, row = rx
            #print finding;
            if finding.lower() == "n/a" or finding.strip() == '':
                continue
        
            findingdetail1 = sheet.cell_value(rowx=rx,colx=5) #Column F, row=rx
            findingdetail2 = sheet.cell_value(rowx=rx,colx=6) #Column G, row=rx
            id = sheet.cell_value(rowx=rx,colx=1) #Column B, row=rx
            cat = sheet.cell_value(rowx=rx,colx=2), #Column C, row=rx
            cause = sheet.cell_value(rowx=rx,colx=8), #Column I

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
        (fdhdr,numbered_finding_details,fdob) = AuditParser(d1+'\n\n'+d2).parse(fob)

        exd_observations = []
        exd_observations.append(fdhdr)
        exd_observations.extend(numbered_finding_details)

        #Check for numeric matches - code matched as Cond3
        if( numbered_findings and numbered_finding_details and len(numbered_findings) == len(numbered_finding_details)):
            numbered_findings[0] = fhdr + "\n" + numbered_findings[0]
            numbered_finding_details[0] = fdhdr + "\n" + numbered_finding_details[0]
            for i in range(len(numbered_findings)):
                ar.write_row(rx+1,id,cat,FINDING_TYPE, numbered_findings[i] , numbered_finding_details[i],cause,"Cond3: Numeric match")
        elif numbered_findings and (not numbered_finding_details or len(numbered_findings) != len(numbered_finding_details)):
            numbered_findings[0] = fhdr + "\n" + numbered_findings[0]
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

    try:
        converter = XLConverter(book,ar.get_workbook(),config_file)
        converter.process()
        
        #get facility profile name from file
        (frow,fcol) = converter.decode_single_cell('B5')
        facility_name = facility_profile.cell_value(rowx=frow,colx=fcol).strip()
        facility_id = facility_mapper.get_facility_id(facility_name,2013)
        converter.write_value("Facility Profile","A1","Facility ID From MSS:")
        converter.write_value("Facility Profile","B1",facility_id)

        ar.save()
        logging.info('Written file: ' + ifile + ".xls")
    except:
        logging.error("Error processing file: " + ifile + ".No output written")
        
