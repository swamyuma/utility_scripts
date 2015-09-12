#!/usr/bin/env python
# Reference: http://stackoverflow.com/questions/5705588/python-creating-excel-workbook-and-dumping-csv-files-as-worksheets
# Date: 2012-08-31, added argparse
# Project: compile csv/dat files as excel worksheets into compiled.xls

# imports
import click
import glob, csv, xlwt, os

# create a workbook object
wb = xlwt.Workbook()

CONTEXT_SETTINGS = dict(help_option_names=['-h', '--help'])

# if a glob option is chosen
@click.group(CONTEXT_SETTINGS)
@click.version_option(version='1.0.0')
def combine():
    pass

def glober(**kwargs):
    output = "{0} {1}".format(kwargs['glob'], kwargs['delimiter'])
    print output  

@combine.command()
@click.option('--delimiter', default=",", help="allowed delimiters are , and \t")
@click.option('--glob', help="enter *.csv")
@click.option('--output', default="combined.xls", help="enter output file name")
def xls_glob(**kwargs):
    '''
    combines all csv files in specified directory
    ''' 
    glober(**kwargs)    

    try:
    # read in the kwargs
        file_glob = kwargs['glob']
        delimiter = kwargs['delimiter'].__str__()
        print '[GLOB] Successfully read in ...', file_glob
    except IOError as e: 
        print '[ERROR] Cannot read file: ', file_glob
        sys.exit(0)

    # foreach file, create a worksheet by
    for filename in glob.glob(file_glob):
        print '[FILENAME] is {0}'.format(filename)
        # calling the save_output function
        save_output(filename, delimiter)
        # save output file  
        try:
            xls_f = kwargs['output']
            print '[SAVING] Writing {0} to {1}'.format(filename, xls_f)
            wb.save(xls_f)
        except IOError as e:
            print '[ERROR] Cannot open xls file for writing:{0}'.format(xls_f)
            sys.exit(0)

## if a list option is chosen
@combine.command()
@click.option('--delimiter', default=",", help="allowed delimiters are , and \t")
@click.option('--list', help="enter *.csv")
@click.option('--output', default="combined.xls", help="enter output file name")
def xls_list(**kwargs):
    '''
    combines csv files specified in a file_list text file.
    ''' 
    # file_list.txt
    try:
        # get the kwargs
        file_list = kwargs['list']
        delimiter = kwargs['delimiter'].__str__()
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
            xls_f = kwargs['output']
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

if __name__=="__main__":
    combine()
