import xlrd
from xlwt import *
import glob
from xlrd import XLRDError

#def removeNumerics(line):
#    if(line.strip()[0].isdigit()):
#        

ROW = 0
FINDING = 1
FINDING_DETAIL = 2
COMMENT = 3

write_rw = 0

def write_row(ws,row_num, finding, detail, comment):
    global write_rw
    ws.write(write_rw,ROW,row_num)
    ws.write(write_rw,FINDING,finding)
    ws.write(write_rw,FINDING_DETAIL,detail)
    ws.write(write_rw,COMMENT,comment)
    write_rw +=1;

#define writer
wb = Workbook()
ws0 = wb.add_sheet('0')

write_row(ws0,"Row #","Finding","Finding Detail", "Comment")


ifiles = glob.glob("/Users/neelanshadwivedi/Desktop/janu/*.xlsx")
for ifile in ifiles:

#print "The number of worksheet is: %d" % book.nsheets
    try:
        book = xlrd.open_workbook(ifile,encoding_override="cp1252")
        sheet = book.sheet_by_name("Findings & CAP")
    except XLRDError:
        print "Unable to open -Findings & CAP- in file: " + ifile
        continue

    for rx in range(2,sheet.nrows):
        finding = sheet.cell_value(rowx=rx,colx=4) #Column E, row = rx
        #print finding;
        if finding.lower() == "n/a":
            continue

        findings = finding.splitlines();
        #print findings

        findingdetail1 = sheet.cell_value(rowx=rx,colx=5) #Column F, row=rx
        findingdetail2 = sheet.cell_value(rowx=rx,colx=6) #Column G, row=rx

        d1 = findingdetail1.strip()  #remove leading whitespace
        d2 = findingdetail2.strip()  #remove leading whitespace

        if d1.lower() == 'n/a' or d1.strip() == '' : d1 = ''
        if d2.lower() == 'n/a'or d2.strip() == '' : d2 = ''

        if len(findings) == 2 and d1 != '' and d2 != '':
        #print each line as it is
            print u"Row:%d:cond2:" % (rx) + findings[0] + "\t" + d1 
            print u"Row:%d:cond2:" % (rx) + findings[1] + "\t" + d2
            write_row(ws0,rx+1,findings[0],d1,"Cond2")
            write_row(ws0,rx+1,findings[1],d2,"Cond2")
        elif len(findings) == 1:
            print u"Row:%d:cond1:" % (rx) + findings[0] + "\t" + d1 + d2;
            write_row(ws0,rx+1,findings[0],d1+d2,"Cond1")

        else:
            findingdetails = (findingdetail1 + "\r\n" + findingdetail2).splitlines()
        
            if(len(findings) != len(findingdetails)):
                print u"Neelansha Madam. Please look at row: %d "  % rx
                write_row(ws0,rx+1,finding,(findingdetail1 + "\r\n" + findingdetail2),"Cond4: Neelansha Please look at it!")
            #TODO: write this on the excel sheet. part complete!
            else:
                for i in range(len(findings)):
                    print u"Row:%d:cond3:" % (rx) + findings[i] + "\t" + findingdetails[i]
                    write_row(ws0,rx+1,findings[i],findingdetails[i],"Cond3")

        wb.save(ifile + ".xls");            



        
    

        
        
        

    
    
    
        
        
        
        


     
    

