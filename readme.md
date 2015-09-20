#csvs_to_xls
A utility script using `click` instead of `argparse` to combine multiple csv files into one excel file with each csv file as a worksheet.  I followed the example shown in this wonderful article [parsing libraries](https://realpython.com/blog/python/comparing-python-command-line-parsing-libraries-argparse-docopt-click/) .

## Requirements
Install the following module using `pip`
- `xlwt`

## Usage
Run the following commands to generate a combined excel worksheet.

### Command for combining files in a file list text file
#### Here the list of files to be combined needs to be in a text file, for e.g.
```sh
cat file_list.txt
file1.csv
file2.csv
```

```sh
python csvs_to_xls.py xls_list --list "file_list.txt" --output "combined_list2.xls"
```
### Command for combining all files with an extension csv
#### Make sure the globing is enclosed in double quotes
```sh
python csvs_to_xls.py xls_glob --glob "*.csv" --output "combined_glob2.xls"
```

# csv_to_xls
A utility script to combine multiple csv files into one excel file with each csv file as a worksheet.

## Requirements
Install the following module using `pip`
- `xlwt`

## Usage
Run the following commands to generate a combined excel worksheet.

### Command for combining files in a file list text file
#### Here the list of files to be combined needs to be in a text file, for e.g.
```sh
cat file_list.txt
file1.csv
file2.csv
```

`python csv_to_xls.py list --file_list file_list.txt -o combined_list.xls`

### Command for combining all files with an extension csv
#### Make sure the globing is enclosed in double quotes
`python csv_to_xls.py glob --glob "*.csv" -o combined_glob.xls`
