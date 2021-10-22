# Word2Excel

Converts test cases written using the ERIGrid HTD Template (https://github.com/ERIGrid2/Holistic-Test-Description) from Word to Excel.

## Usage

```
usage: word2excel.py [-h] [-t EXCEL_TEMPLATE] [-f] [-c] path

Converts test cases according to the ERIGrid HTD Template from Word into Excel files.

positional arguments:
  path                  Path to either a Word file or a folder. If a folder is provided, all Word files in that folder will be converted.

optional arguments:
  -h, --help            show this help message and exit
  -t EXCEL_TEMPLATE, --excel-template EXCEL_TEMPLATE
                        Path to the Excel template that should be used. Standard: ./template/HTD_TEMPLATE_V1.2.xlsx
  -f, --create-folder   Saves the Excel file and extracted images to a folder with the name of Word file.
  -c, --copy-word-file  Copies the Word file into the new folder
```