import xlwt
from xlutils.copy import copy
import xlrd
import glob
import sys
import logging


class BatchFileProcessor(object):

    def __init__(self,pattern,out_suffix):
        self.ifiles = glob.glob(pattern)
        self.out_suffix = out_suffix

    def process(self):
        for ifile in self.ifiles:
            logging.info("Processing file:" + ifile)
            try:
                inbook = xlrd.open_workbook(ifile)
                outbook = copy(inbook)
            except xlrd.XLRDError as e:
                logging.exception("Unable to open workbook" + ifile + ":" + str(e))
                continue
            
            try:
                if not self.transform(inbook,outbook):
                    logging.error("Unable to process workbook:" + ifile + ":")
                    continue
            except xlrd.XLRDError as e:
                logging.exception("Unable to process file: " + ifile + ":" + str(e))
                continue
            
            out_file = ifile + "." + self.out_suffix
            outbook.save(out_file)
            logging.info('Written file: ' + out_file )
            

    def transform(self,inbook,outbook):
        return True
