#!/usr/bin/python

# Rename:
#Takes an optional rename file in the format:
#FileA::FileB
#..
# and renames the file

import os
import getopt
import sys
import hashlib
import glob

FILE_PREFIX = 5
increment_integer = 0
def generate_file_name(basename):
    global increment_integer,FILE_PREFIX
    increment_integer += 1
    reasonable_basename = ''.join(e for e in basename if e.isalnum()).lower()
    return reasonable_basename[0:FILE_PREFIX] + str(increment_integer)

optlist,gargs = getopt.getopt(sys.argv[1:],'c:s:e:d:g:')
optmap = dict(optlist)

map_file = optmap.get("-c")
sym_link = optmap.get("-s")
extension = optmap.get("-e","")
directory = optmap.get("-d","")
fileglob = optmap.get("-g","")

print directory

#files = list(gargs)
#if gargs and fileglob:
files  = glob.glob(fileglob)

#print optmap

#print files
#print map_file

rename_map = {}
if map_file:
    with file(map_file) as f: file_contents = f.read()
    rename_map = dict( line.split('::') for line in file_contents)

print files
for fi in files:
    print fi
    if os.path.isfile(fi):
        tgt_folder = directory if directory and os.path.dirname(directory) else "."
        if rename_map:
            tgt_file = os.path.dirname(fi)+"/"+(rename_map[fi])
        else:
            tgt_file = tgt_folder + "/" + generate_file_name(os.path.basename(fi)) #hashlib.md5(fi.encode('utf-8')).hexdigest() + "." + extension

        print tgt_file
        os.symlink(fi,tgt_file) if sym_link else os.rename(fi,tgt_file)
        print( tgt_file + '::' + os.path.basename(fi))
