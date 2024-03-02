# User List CSV Tool

This Python script is designed to process Excel files and generate two separate CSV files, one containing newly added users and another containing retired users.

## Usage

```bash
python3 user-list-csv-tool.py <.xlsx file> <old sheet roster name> <new sheet roster name>
```


## Prerequisites for Using `user-list-csv-tool.py`

Before using the `user-list-csv-tool.py` script, install the required dependencies using the provided setup script.

For Linux-based systems, use the following script to install the required Python libraries:
(uses `apt` package manager)
```bash
./linux-setup-python-libraries.sh
```

For Windows systems, use the following script to install the required Python libraries (RUN AS ADMINISTRATOR)
(uses `choco` package manager)
```bash
./windows-setup-python-libraries.sh
```

## Dependencies Installed by setup-python-libraries.sh

The setup script, ```setup-python-libraries.sh```, will install the following dependencies:

    pip (Python library installer)
    pandas (data manipulation library for Python)
    openpyxl (library for reading and writing Excel files)

**Note:** This script is designed to work on Linux-based systems with the `apt` package manager. If you are using Windows, it is recommended to install a package manager such as [Chocolatey (choco)](https://chocolatey.org/) or another Windows-based package manager and modify the script accordingly

## Important information for input .xlsx file

