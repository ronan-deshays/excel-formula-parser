# excel-formula-parser
 Standalone python3 script which enable multiline, versionning and comment for Excel Formulas.
 
## Prerequisites
Python 3.10 and above - [install](https://www.python.org/downloads/)
Excel 2019 and above - [install](https://www.microsoft.com/microsoft-365/excel)
 
## Installation
* clone repo locally
* launch python in working folder

## Features
enable Excel Formula :
*  versionning formulas using git
*  writing a formula on multiple lines
*  adding comments into Excel formula

## Parsing process description
* python-like comments management :
 comment is recognized as a line which begin with "# " (hastag followed by space
 
* line breaks and spaces remove

* create one line per formula :
 formula is detected as a line break followed by "=" (equal) sign

## Usage example
* add your formula to [excel_formula_in.txt](https://github.com/ronan-deshays/excel-formula-parser/blob/master/samples/excel_formula_in.txt)
* run [parser.py](https://github.com/ronan-deshays/excel-formula-parser/blob/master/samples/excel_formula_in.txt)
* get result in [excel_formula_out.txt](https://github.com/ronan-deshays/excel-formula-parser/blob/master/samples/excel_formula_in.txt)

## Limitations
* variable name containing space not supported
* unable to make difference between formula begin and "=" sign in formula body

## Perspectives
* Excel linting needs to copy formula to Excel
* enable multiline comments
* add a sub-functions copy ability (like what a python interpreter does)
