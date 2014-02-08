import re
import sys
import utils


f = open(sys.argv[1])
s = f.read()
print(s.strip())
numerics = utils.extract_numbered_entries(s)
print(numerics)
print(len(numerics))


