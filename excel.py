import sys
import re
import xlrd
from xlwt import *
import glob
from xlrd import XLRDError

#def removeNumerics(line):
#    if(line.strip()[0].isdigit()):
#

#TODO: define a class which is the AuditReportWriter
# It defines all the constants
# It keeps track of new workbook, row number, column number etc..
# This will be the only dependency with xlwt

ROW = 0
ID = 1
FINDING = 2
FINDING_DETAIL = 3
COMMENT = 4

write_rw = 0

def write_row(ws,row_num, id,finding, detail, comment):
    global write_rw

    ws.write(write_rw,ROW,row_num)
    ws.write(write_rw,ID,id)
    ws.write(write_rw,FINDING,finding)
    ws.write(write_rw,FINDING_DETAIL,detail)
    ws.write(write_rw,COMMENT,comment)
    write_rw +=1

def extract_numbered_entries(blob):
    lines = blob.splitlines()
    numbered_list = []
    numbered = None
    buf = ''

    for line in lines:
        if re.match('^\d\s*[.):].*',line): #beginning of a new numbered list
            if buf and numbered: #if its not the first numbered list
                numbered_list.append(buf) #push the last numbered list
                buf = '' #empty buffer
            numbered = True
        #continue buffering
        buf += (line +'\n')

    #closing block
    if buf and numbered: #if there is a buffer and its a numbered list
        numbered_list.append(buf) #push the last numbered list
        buf = ''

    return numbered_list


ifiles = glob.glob(sys.argv[1])
for ifile in ifiles:
    print "Processing file:" + ifile
#print "The number of worksheet is: %d" % book.nsheets
    try:
        book = xlrd.open_workbook(ifile,encoding_override="cp1252")
        sheet = book.sheet_by_name("Findings & CAP")
    except XLRDError:
        print "Unable to open -Findings & CAP- in file: " + ifile
        continue

    #define writer
    #TOD0: encapsulate the creation and initialization in to a method in a class
    # leads to better abstraction of reads and writes. Easier to debug!!
    write_rw = 0  
    wb = Workbook()
    ws = wb.add_sheet('0')
    write_row(ws,"Row #","ID", "Finding","Finding Detail", "Comment")

    for rx in range(2,sheet.nrows):
        finding = sheet.cell_value(rowx=rx,colx=4) #Column E, row = rx
        #print finding;
        if finding.lower() == "n/a" or finding.strip() == '':
            continue

        findingdetail1 = sheet.cell_value(rowx=rx,colx=5) #Column F, row=rx
        findingdetail2 = sheet.cell_value(rowx=rx,colx=6) #Column G, row=rx
        id = sheet.cell_value(rowx=rx,colx=1) #Column A, row=rx

        try:
            findings = finding.splitlines();

            d1 = findingdetail1.strip()  #remove leading whitespace
            d2 = findingdetail2.strip()  #remove leading whitespace

            if d1.lower() == 'n/a' or d1.strip() == '' : d1 = ''
            if d2.lower() == 'n/a'or d2.strip() == '' : d2 = ''
        except:
            print "Error reading, row: " + str(rx) + ":" + str(sys.exc_info())
            #try to write row back
            write_row(ws,rx+1,id,finding,unicode(findingdetail1)+"\n" + unicode(findingdetail2),"Cond4: Neelansha Madam. Please look" + str(sys.exc_info()))
            print(unicode(findingdetail1)+"\n" + unicode(findingdetail2))
            continue
            

        numbered_findings = extract_numbered_entries("\n".join(findings))
        numbered_finding_details = extract_numbered_entries(d1+'\n'+d2)

        #Check for numeric matches - code matched as Cond3
        if( numbered_findings and numbered_finding_details and len(numbered_findings) == len(numbered_finding_details)):
            for i in range(len(numbered_findings)):
                #print u"Row:%d:cond3 numeric:" % (rx) + numbered_findings[i] + "\t" + numbered_finding_details[i]
                write_row(ws,rx+1,id,numbered_findings[i],numbered_finding_details[i],"Cond3: Numeric match")
        #Cond2 - two findings with d1 and d2
        elif not numbered_findings:
            #equality check for 2 - code named Cond2
            if len(findings) == 2 and d1 != '' and d2 != '':
                write_row(ws,rx+1,id,findings[0],d1,"Cond2")
                write_row(ws,rx+1,id,findings[1],d2,"Cond2")
            #one finding with d1 and optional d2 - code named Cond1
            elif len(findings) == 1:
                #print u"Row:%d:cond1:" % (rx) + findings[0] + "\t" + d1 + d2;
                write_row(ws,rx+1,id,findings[0],d1+d2,"Cond1")
            #Cond4 - unmatched findings and finding details after line split
            else:
                write_row(ws,rx+1,id,finding,(d1 + "\n" + d2),"Cond4: Neelansha Please look at it!")
        # Neelansha madam's special case - Cond2.1
        elif numbered_findings and not numbered_finding_details and len(numbered_findings) == 2 and d1 != '' and d2 != '':
            write_row(ws,rx+1,id,numbered_findings[0],d1,"Cond2.1")
            write_row(ws,rx+1,id,numbered_findings[1],d2,"Cond2.1")
        else:
            #Cond4 - no such match. Dump findings and finding details and notify
            write_row(ws,rx+1,id,finding,(d1 + "\n" + d2),"Cond4: Neelansha Please look at it!")


    print 'Written file: ' + ifile + ".xls"    
    wb.save(ifile + ".xls")
