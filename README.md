# User List CSV Tool
## Prerequisites for using `user-list-csv-tool.py`

Before using the `user-list-csv-tool.py` script, install the required dependencies using the provided setup script(s).

For <u>Linux</u>, use the following script to install the required Python libraries:
(uses `apt` package manager)
```bash
./linux-setup-python-libraries.sh
```

For <u>Windows</u>, use the following script to install the required Python libraries (RUN AS ADMINISTRATOR)

This script uses `choco` package manager - https://www.youtube.com/watch?v=-5WLKu_J_AE (link to install `choco`)
```bash
./windows-setup-python-libraries.sh
```

## Dependencies Installed by setup script(s)
The setup script(s), ```windwows-setup-python-libraries.sh / linux-setup-python-libraries.sh```, will install the following dependencies:

    pip (Python library installer)
    pandas (data manipulation library for Python)
    openpyxl (library for reading and writing Excel files)

## Terminology Importance (READ PLEASE!!!)
It is imperative that you understand the difference between a `.xlsx file`, and an  `excel sheet`

A `.xlsx file` is an Excel Workbook that holds 1 or more `Excel sheets`

`Excel sheets` are held within a `.xlsx file`


## Using `user-list-csv-tool.py`
### What does `user-list-csv-tool.py` do?

The Python script is designed to <u>process Excel files and generate two separate CSV files</u>, one containing newly added users and another containing retired users.


### Why use Python?
The reason we are using python is for its **readablility** and **ease of use**. When dealing with excel (.xlsx), it can become extremely cumbersome dealing with sheet reading and parsing. Python simplifies this with libraries, specifically the `pandas` library. Furthermore, whenever we are automating a task, it is critical that the code is readable so that others understand what the script is doing.

### Usage/Input
```bash
python3 user-list-csv-tool.py <.xlsx file> <old sheet roster name> <new sheet roster name>
```

`<.xlsx file>` the file that holds the SHEETS of the rosters being compared

`<old sheet roster name>` the sheet name in the .xlsx file that holds the roster from the previous semester

`<new sheet roster name>` the sheet name in the .xlsx that holds the roster from the current semester

### Output
The program will generate 2 separate .csv files: 
* `users_new.csv` - holds the users that have been added to the "new" excel sheet with users

* `users_retired.csv` - holds the users that were in the "old" excel sheet but are no longer the in "new" excel sheet 







