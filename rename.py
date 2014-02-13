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


optlist,gargs = getopt.getopt(sys.argv[1:],'c:s:e:d:')
optmap = dict(optlist)

map_file = optmap.get("-c")
sym_link = optmap.get("-s")
extension = optmap.get("-e","")
directory = optmap.get("-d","")

#print optmap

files = gargs
#print files
#print map_file

rename_map = {}
if map_file:
    with file(map_file) as f: file_contents = f.read()
    rename_map = dict( line.split('::') for line in file_contents)

for fi in files:
    if os.path.isfile(fi):
        tgt_folder = os.path.dirname(fi) if not directory else "."
        if rename_map:
            tgt_file = os.path.dirname(fi)+"/"+(rename_map[fi])
        else:
            tgt_file = tgt_folder +"/" + hashlib.md5(fi.encode('utf-8')).hexdigest() + "." + extension

        print tgt_file
        os.symlink(fi,tgt_file) if sym_link else os.rename(fi,tgt_file)
        print( tgt_file + '::' + os.path.basename(fi))
