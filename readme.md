# csv_to_xls
A utility script to combine multiple csv files into one excel file with each csv file as a worksheet.

## Requirements
Install the following module using `pip`
- xlwt

## Usage
Run the following commands to generate a combined excel worksheet.

### command for combining files in a file list text file
#### Here the list of files to be combined needs to be in a text file, for e.g.
```sh
cat file_list.txt
file1.csv
file2.csv
```

`python csv_to_xls.py list --file_list file_list.txt -o combined_list.xls`

### command for combining all files with an extension csv
#### Make sure the globing is enclosed in double quotes
`python csv_to_xls.py glob --glob "*.csv" -o combined_glob.xls`
