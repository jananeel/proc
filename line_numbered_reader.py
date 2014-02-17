import re
import sys
from audit_parser import AuditParser 
import utils


f = open(sys.argv[1])
s = f.read()
print(s.strip())
numerics = utils.extract_numbered_entries(s)
print(numerics)
print(len(numerics))
all = AuditParser(s).parse()
print(all)
print(len(all[1]))




