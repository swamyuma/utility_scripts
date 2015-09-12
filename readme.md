# csv_to_xls
A utility script to combine multiple csv files into one excel file with each csv file in a worksheet.

# Requirements
python modules
- xlwt

# Usage
Run the following commands to generate a combined excel worksheet.

## command for combining files in a file list text file
`python csv_to_xls.py list --file_list file_list.txt -o combined_list.xls`

## command for combining all files with an extension csv
`python csv_to_xls.py glob --glob "*.csv" -o combined_glob.xls`
