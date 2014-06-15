import pyodbc, csv, sys, os

connect_string = 'DRIVER={Microsoft Access Driver (*.mdb)};DBQ=Z:\\Jo\\downloads\\Lancaster_uad.mdb'

def get_tables(connection):
    try:
        cur = connection.cursor()
        return [row[2] for row in (s for s in cur.tables() if s[3] != 'SYSTEM TABLE')]
    except:
        e = sys.exc_info()[1]
        print "Error1 %s" % e

def get_data(tblName, connection):
    try:
        cursor = connection.cursor()
        cursor.execute('SELECT * FROM %s' %(tblName))
        return [row for row in cursor]
    except:
        print "There is a problem with getting data from %s" % tblName
        pass #fail on this one, skip and move on

def get_columns(tblName, connection):
    try:
        cursor = connection.cursor()
        return [row[3] for row in cursor.columns(table = tblName)]
    except:
        print "There is a problem with getting columns from %s" % tblName
        pass

def qexport():
    connection = pyodbc.connect(connect_string)
    for i in get_tables(connection):
        try:
            filename = i + '.csv'
            outfile = open(filename, 'wb') 
            writer = csv.writer(outfile)
            writer.writerow(get_columns(i, connection))
            writer.writerows(get_data(i, connection))
            outfile.close()
        except:
            outfile.close()
            os.remove(filename) #remove the csv we've just created
            print "There is a problem with %s" % i
            pass

if __name__ == "__main__":
    qexport()
