import argparse

parser = argparse.ArgumentParser(description='Read xls file into db')
parser.add_argument('filename', help='filename of xls file')
parser.add_argument('--database', help='database to use (default: basename of the file)')
parser.add_argument('--sheet',metavar='N',type=int, default=1, help='sheet number (default: 1)')
parser.add_argument('--header',metavar='N', type=int, default=1, help='headerline (default: 1)')
parser.add_argument("-l","--lines",metavar='LIST', help="list of lines and line ranges")

# column types: int float double char64 text time date dt=datetime
parser.add_argument("-i","--int",metavar='LIST', help="columns of integer type")
parser.add_argument("-f","--float",metavar='LIST',help="columns of float type (double precision)")
parser.add_argument("-c","--char",metavar='LIST',help="columns of character type")
parser.add_argument("--time",metavar='LIST',help="columns of time type")
parser.add_argument("--date",metavar='LIST',help="columns of date type")
parser.add_argument("--datetime",metavar='LIST',help="columns of datetime type")

parser.add_argument("-v", "--verbose", help="verbose diagnostic output", action="store_true")
parser.add_argument("--drop", help="drop any existing tables", action="store_true")

args = parser.parse_args()


# function for printing to STDERR in verbose mode
import sys
def eprint(*arguments, **kwargs):
    if args.verbose:
        print(*arguments, file=sys.stderr, **kwargs)

# select database
import os
DATABASE=os.path.splitext(os.path.basename(args.filename))[0]
if args.database: DATABASE=args.database
eprint("DATABASE:")
eprint(DATABASE)
eprint()

# process column type selection
def getcols(attribute):
    s=getattr(args,attribute)
    cols=[]
    if s:
        cols=list(map(int,s.split(",")))
    return cols

intcols=getcols("int")
floatcols=getcols("float")
textcols=getcols("char")

import xlrd
book = xlrd.open_workbook(args.filename)


eprint("SHEET:")
sh = book.sheet_by_index(args.sheet - 1)
SHEETNAME=book.sheet_names()[args.sheet - 1]
eprint(SHEETNAME)
eprint()

xlrdtypes =["XL_CELL_EMPTY",
	    "XL_CELL_TEXT",
	    "XL_CELL_NUMBER",
	    "XL_CELL_DATE",
	    "XL_CELL_BOOLEAN",
	    "XL_CELL_ERROR",
	    "XL_CELL_BLANK"]

eprint("HEADER:")
h=[]
for cx in range (sh.ncols):
    cell = sh.cell(args.header-1,cx)
    h.append(cell.value)
eprint(h)
eprint()

# initialize list of types mapped to the fields
eprint("TYPES:")
fieldlist=['']*sh.ncols
for i in intcols: fieldlist[i-1]="INT"
for i in floatcols: fieldlist[i-1]="DOUBLE"
for i in textcols: fieldlist[i-1]="TEXT"
eprint(fieldlist)
eprint()

# print list of cells in verbose mode
eprint("CELLS:")
if args.verbose:
    for rx in range(args.header, sh.nrows):
        for cx in range (sh.ncols):
            cell = sh.cell(rx,cx)
            eprint("{0} {1} {2:7} ".format(rx,cx,fieldlist[cx]), end='')
            eprint(cell.value)
    eprint()


#######   START OF SQL OUTPUT

# preparation
print ("USE "+ DATABASE + ";")
if args.drop: print("DROP TABLE IF EXISTS "+ SHEETNAME + ";")
print ("CREATE TABLE " + SHEETNAME + " (",end='')
comma=False
for i in range(len(fieldlist)):
    if fieldlist[i] != "":
        if comma: print(', ',end='')
        print ("{0} {1}".format(h[i],fieldlist[i]), end='')
        comma=True
print (");")   

# formatting a value to insert format
def printvalue(x,fieldtype):
    if fieldtype=="INT":
        print("{0:d}".format(int(x)),end='')
    elif fieldtype=="DOUBLE":
        print('{0}'.format(x),end='')
    elif fieldtype=="TEXT":
        print('"{0}"'.format(x),end='')

# convert one row of input to a record for inserting
def processrow(rx):
    print ("INSERT INTO " + SHEETNAME + " (",end='')
    comma=False
    for i in range(len(fieldlist)):
        if fieldlist[i] != "":
            if comma: print(',',end='')
            print(h[i],end='')
            comma=True
    print (") VALUES (",end='')
    comma=False
    for i in range(len(fieldlist)):
        if fieldlist[i] != "":
            if comma: print(',',end='')
            cell=sh.cell(rx,i)
            printvalue(cell.value,fieldlist[i])
            comma=True
    print(");")

# process cells row by row
        
if args.lines:
    for entry in args.lines.split(','):
        if '-' in entry:
            rangeentry=entry.split('-')
            start=int(rangeentry[0])-1
            end=int(rangeentry[1])
            for rx in range(start, end):
                processrow(rx)
        else:
            processrow(int(entry)-1)
else:
    for rx in range(args.header, sh.nrows):
        processrow(rx)
    
