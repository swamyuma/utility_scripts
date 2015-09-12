#!/usr/bin/env python
# Reference: http://stackoverflow.com/questions/5705588/python-creating-excel-workbook-and-dumping-csv-files-as-worksheets
# Date: 2012-08-31, added argparse
# Project: compile csv/dat files as excel worksheets into compiled.xls

# imports
import argparse
import glob, csv, xlwt, os

# create a workbook object
wb = xlwt.Workbook()


# if a glob option is chosen
def xls_glob(args):
    # print input arguments in the current Namesapce
    #print '[ARGS]...', args.glob[0]
    # read inputs
    #for filename in glob.glob("c:/MyPython/A*.33.csv"):
    try:
    # read in the first argument as a string
        file_glob = args.glob[0]
        delimiter = args.DELIM[0]
        print '[GLOB] Successfully read in ...', file_glob
    except IOError as e: 
        print '[ERROR] Cannot read file: ', file_glob
        sys.exit(0)

    # foreach file, create a worksheet by
    for filename in glob.glob(file_glob):
        # calling the save_output function
        save_output(filename, delimiter)
        # save output file  
        try:
            xls_f = args.output
            print '[SAVING] Writing {0} to {1}'.format(filename, xls_f)
            wb.save(xls_f)
        except IOError as e:
            print '[ERROR] Cannot open xls file for writing:{0}'.format(xls_f)
            sys.exit(0)

# if a list option is chosen
def xls_list(args):
    # print input arguments in the current Namesapce
    print args
    # read inputs
    # file_list.txt
    try:
        # read in the first argument as a string
        file_list = args.file_list[0]
        delimiter = args.DELIM[0]
        print '[INIT] Successfully read in ...', file_list
    except IOError as e: 
        print '[ERROR] Cannot read file: ',  file_list
        sys.exit(0)
    # use a list comprehension for reading lines in the file list   
    lines = [lines.strip('\n') for lines in open(file_list)]
    # foreach file, create a worksheet by   
    for filename in lines:
        # calling the save_output function  
        save_output(filename, delimiter)
        # save output file          
        try:
            xls_f = args.output
            print '[SAVING] Writing {0} to {1}'.format(filename,xls_f)
            wb.save(xls_f)

        except IOError as e:
            print '[ERROR] Cannot open xls file for writing:{0}'.format(xls_f)
            sys.exit(0)

# function for writing to worksheets
def save_output(inputfile, delimiter):
    filename = inputfile
    this_delimiter = delimiter
    (f_path, f_name) = os.path.split(filename)
    (f_short_name, f_extension) = os.path.splitext(f_name)
    # create worksheet name
    ws = wb.add_sheet(f_short_name)
    # open worksheet to write 
    spamReader = csv.reader(open(filename, 'rb'), delimiter=this_delimiter)
    for rowx, row in enumerate(spamReader):
        for colx, value in enumerate(row):
            ws.write(rowx, colx, value)


# version information   
VERSION = "0.1.2"

# create a parser object with version information
parser = argparse.ArgumentParser(prog="CSV_TO_XLS", description="Compile text files into excel worksheets %s" % VERSION)
parser.add_argument('--version', action='version', version='%(prog)s ' + VERSION)

# create a subparser argument
subparsers = parser.add_subparsers(help='Command to be run')

# file_list
file_parser = subparsers.add_parser('list', help='Compile from a file list')
file_parser.add_argument('--file_list', action='store', required=True, metavar='/path/to/file_list', help='File list for compiling', nargs=1)
file_parser.add_argument('--input_delimiter', action='store', dest="DELIM", default=",", help='specify input delimiter, default(,)', type=str)
file_parser.add_argument('--output', '-o', action='store', required=True, metavar='/path/to/output.xls', help='Output location of the compiled xls', type=str)

# call the glob function and create output
file_parser.set_defaults(func=xls_list)

#file_glob
glob_parser = subparsers.add_parser('glob', help='Compile from a glob')
glob_parser.add_argument('--glob', action='store', required = True, metavar='"/path/to/*.csv"', help='enlcose glob within double quotes', nargs=1 )

glob_parser.add_argument('--input_delimiter', action='store', dest="DELIM", default=",", help='specify input delimiter, default(,). Escape sequences like \\t are allowed.', type=str)

glob_parser.add_argument('--output', '-o', action='store', required=True, metavar='/path/to/output.xls', help='Output location of the compiled xls', type=str)

# call the list function and create output
glob_parser.set_defaults(func=xls_glob)

# parse arguments
args = parser.parse_args()
# decode escape sequences
args.DELIM = args.DELIM.decode('string_escape')
args.func(args)
