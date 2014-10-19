# -*- coding: utf-8 -*-
#!/usr/bin/env python
#
# AccessDump.py
# A simple script to dump the contents of a Microsoft Access Database. Initially
# copied from http://mazamascience.com/WorkingWithData/?p=168 with some minor changes
# Thanks Guys!
# It depends upon the mdbtools suite:
#   http://sourceforge.net/projects/mdbtools/

import sys, subprocess, os

DATABASE = sys.argv[1]

# Get the list of table names with "mdb-tables"
table_names = subprocess.Popen(["mdb-tables", "-1", DATABASE], 
                               stdout=subprocess.PIPE).communicate()[0]
tables = table_names.split('\n')

# Create an output directory below the current one for tidiness
# Pass if it already exists
outputdir = os.path.join(os.getcwd(), 'output')
try:
	os.makedirs(outputdir, 0755)
except OSError:
	pass

# Dump each table as a CSV file uinto output directory using "mdb-export",
# converting " " in table names to "_" for the CSV filenames.
for table in tables:
    if table != '':
        filename = os.path.join(outputdir, table.replace(" ","_") + ".csv")
        file = open(filename, 'w')
        print("Dumping " + table)
        contents = subprocess.Popen(["mdb-export", DATABASE, table],
                                    stdout=subprocess.PIPE).communicate()[0]
        file.write(contents)
        file.close()