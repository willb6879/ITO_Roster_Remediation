import pandas as pd
import secrets
import string
import sys
import os
import subprocess as sp

RED_TEXT = "\033[91m"
CSV_USERS_NEW = "users_new.csv"
CSV_USERS_RETIRED = "users_retired.csv"
PS_SCRIPT = "gen-ad-users.ps1"
ERR_INVALID_ARGS = 1
ERR_FILE_NOT_EXIST = 2
ERR_INVALID_HDR = 3

def print_err(ERROR_CODE, f):
    """
    Prints error message based on an error code

    Parameters:
    - ERROR_CODE (int): The error code input
    - f: optional input to use for error printing description

    Returns: 
    - None: function prints out error message
    """
    if ERROR_CODE == ERR_INVALID_ARGS:
        # Invalid number of arguments
        print(f"{RED_TEXT}Error: invalid number of arguments")
        print(f"{RED_TEXT}Usage: python3 user-list-csv-tool.py <.xslx file> <old sheet roster name> <new sheet roster name>\n")
        print(f'{RED_TEXT}Example...\n\npython3 user-list-csv-tool.py ito-rosters.xlsx "Roster S22" "Roster F22"\n')
        return
    if ERROR_CODE == ERR_INVALID_HDR:
        # Invalid headers in either excel sheet
        print(f"{RED_TEXT}Error: invalid headers in '{f}'")
        print(f"{RED_TEXT}Please ensure that each input excel sheet has the following headers...\n\n")
        print(f"{RED_TEXT}Course, Name, Class, Major, Email, Group\n\n")
        print(f"{RED_TEXT}Hint: You can copy the excel sheet labeled 'Roster Template' to a new sheet to ensure these headers are correct")
        return
    if ERROR_CODE == ERR_FILE_NOT_EXIST:
        print(f"{RED_TEXT}Error: '{f}' does not exist in the current working directory")
        return


def file_exists(file):
    """
    Check if a file exists at the specified path.

    Parameters:
    - file_path (str): The path to the file.

    Returns:
    int: 0 if the file exists, -1 if the file doesn't exist.

    Example:
    >>> file_exists("example.txt")
    0

    >>> file_exists("nonexistent_file.txt")
    -1
    """
    if os.path.exists(file):
        # file exists
        return 0
    else:
        # file doesn't exist
        return -1


# Returns the number of input arguments
def get_num_args():
    """
    Return the number of arguments input by the system.

    Returns:
    int: The number of command-line arguments provided by the system.

    Example:
    >>> get_num_args()
    3

    Note:
    The count includes the script name as the first argument.
    """
    num_args = len(sys.argv) - 1
    return num_args

# Sanity checks the headers in an excel sheet
# Returns 0 on success, -1 on failure
def check_sheet_headers(dataframe):
    """
    Sanity checks the headers in an Excel sheet against expected headers.

    Parameters:
    - dataframe (pandas.DataFrame): The pandas DataFrame representing the Excel sheet.

    Returns:
    - int: Returns 0 on success, indicating that the Excel sheet headers match the expected headers.
            Returns -1 on failure, indicating a mismatch or unexpected error.

    Description:
    - This function performs a sanity check on the headers of an Excel sheet represented by a pandas DataFrame.
    - It compares the actual headers in the DataFrame with a predefined list of expected headers.
    - If the headers match, the function returns 0. If there is a mismatch or an unexpected error occurs, it returns -1.
    """
    # Gets headers from the dataframe
    expected_headers = ["Course", "Name", "Class", "Major", "Email", "Group"]
    df_headers = list(dataframe.columns)

    # Performs sanity check for excel sheet header
    try:
        for i in range(0, len(df_headers)):
            if expected_headers[i] != df_headers[i]:
                # Non-matching header
                return -1
    except IndexError as e:
        # Index out of bounds - doesn't match headers
        return -1
    except Exception as e:
        print(f"{RED_TEXT}Unexpected error has occurred\n {e}")
        sys.exit(1)
    
    # Input excel sheet headers are compliant
    return 0


def gen_rand_password(length=12):
    """
    Generate a random password with Active Directory complexity requirements.

    Parameters:
    - length (int): The length of the password.

    Returns:
    str: A random password.
    """
    uppercase_letters = string.ascii_uppercase
    lowercase_letters = string.ascii_lowercase
    digits = string.digits
    special_characters = string.punctuation

    # Ensure at least one character from each category
    password = (
        secrets.choice(uppercase_letters) +
        secrets.choice(lowercase_letters) +
        secrets.choice(digits) +
        secrets.choice(special_characters)
    )

    # Fill the rest of the password with random characters
    password += ''.join(secrets.choice(uppercase_letters + lowercase_letters + digits + special_characters) for _ in range(length - 4))

    # Shuffle the password to ensure randomness
    password_list = list(password)
    secrets.SystemRandom().shuffle(password_list)
    return ''.join(password_list)

# Formats csv to desired output    
def formatCSV(csv_file, removed_columns):
    """
    Formats a CSV file by adding and modifying specific columns.

    Parameters:
    - csv_file (str): The path to the CSV file to be formatted.
    - removed_columns (list): A list of column names to be removed from the CSV.

    Returns:
    None: The function modifies the CSV file in-place.

    Example:
    >>> formatCSV('input.csv', ['Unnecessary_Column1', 'Unnecessary_Column2'])
    
    Description:
    This function reads a CSV file, adds 'Username', 'Firstname', 'Lastname', and 'TempPassword' columns based on existing columns,
    removes specified unnecessary columns, and then saves the modified DataFrame back to the original CSV file.

    Note:
    - The 'Username', 'Firstname', 'Lastname', and 'TempPassword' columns are added in specific positions.
    - The function modifies the CSV file directly and does not return a new DataFrame.

    """
    # Read in csv file
    df = pd.read_csv(csv_file)

    # Add 'Username', 'Firstname', 'Lastname', and 'TempPassword' Columns
    df['Username'] = df['Email'].str.split('@').str[0]
    username_col = df.pop('Username')
    df.insert(0, 'Username', username_col)

    df['Firstname'] =  df['Name'].str.split(' ').str[0]
    username_col = df.pop('Firstname')
    df.insert(3, 'Firstname', username_col)

    df['Lastname'] =  df['Name'].str.split(' ').str[1]
    username_col = df.pop('Lastname')
    df.insert(4, 'Lastname', username_col)

    # Remove remaining unecessary columns
    df = df.drop(columns=removed_columns)
    df.to_csv(csv_file, index=False)


def gen_user_csv_files(excel_file, old_sheet_roster, new_sheet_roster):
    """
    Generate CSV files containing information about users who have joined and left the enterprise.

    Parameters:
    - excel_file (str): The path to the Excel file containing user rosters.
    - old_sheet_roster (str): The name or index of the sheet in the Excel file representing the old user roster.
    - new_sheet_roster (str): The name or index of the sheet in the Excel file representing the new user roster.

    Returns:
    None: The function saves the results to CSV files.

    Example:
    >>> gen_user_csv_files('user_data.xlsx', 'OldRoster', 'NewRoster')

    Description:
    This function reads two sheets from an Excel file representing old and new user rosters, respectively.
    It then identifies users who have left and joined the enterprise based on changes in the 'Email' column.
    The information about retired users is saved to a CSV file named 'users_retired.csv', and the information
    about new users is saved to a CSV file named 'users_new.csv'. The CSV files are formatted using the 'formatCSV'
    function to remove unnecessary columns such as 'Course', 'Major', 'Class', and 'Email'.

    Note:
    - The CSV files are saved with headers in 'users_new.csv'.
    - The 'Course', 'Major', 'Class', and 'Email' columns are removed from both CSV files during formatting.

    """
    # Read in excel sheets
    try:
        old_df = pd.read_excel(excel_file, sheet_name=old_sheet_roster)
    except ValueError as e:
        print(f"{RED_TEXT}Error: Worksheet '{old_sheet_roster}' not found in '{excel_file}'")
        sys.exit(1)   
    except Exception as e:
        print(f"{RED_TEXT}Unexpected error has occurred\n {e}")
        sys.exit(1)
    try:
        new_df = pd.read_excel(excel_file, sheet_name=new_sheet_roster)
    except ValueError as e:
        print(f"{RED_TEXT}Error: Worksheet '{new_sheet_roster}' not found in '{excel_file}'")
        sys.exit(1)
    except Exception as e:
        print(f"{RED_TEXT}Unexpected error has occurred\n {e}")
        sys.exit(1)

    # Check each sheet headers
    if (check_sheet_headers(old_df) == -1):
        print_err(3, old_sheet_roster)
        sys.exit(1)
    if (check_sheet_headers(new_df) == -1):
        print_err(3, new_sheet_roster)
        sys.exit(1)

    # Determine users that have left the enterprise
    retired_users = old_df[~old_df['Email'].isin(new_df['Email'])]

    # Determine users that have joined the enterprise
    new_users = new_df[~new_df['Email'].isin(old_df['Email'])]

    # Add TempPassword column to new_users dataframe
    new_users = new_users.assign(TempPassword=None)

    # Add temporary passwords to each respective new user in the dataframe
    for index, row in new_users.iterrows():
        new_users.loc[index, 'TempPassword'] = gen_rand_password()

    # Save results to csv files
    new_users.to_csv(CSV_USERS_NEW, index=False, header=True)
    retired_users.to_csv(CSV_USERS_RETIRED, index=False)

    # Format csv files
    formatCSV(CSV_USERS_NEW, ['Course', 'Major', 'Class', 'Email'])
    formatCSV(CSV_USERS_RETIRED, ['Course', 'Major', 'Class', 'Email'])
    

if __name__ == "__main__":
    # Check the number of arguments
    if get_num_args() != 3:
        # Invalid number of arguments
        print_err(ERR_INVALID_ARGS, 0)
        sys.exit(1)
    
    # Check excel file exists
    if file_exists(sys.argv[1]) == -1:
        print_err(ERR_FILE_NOT_EXIST, sys.argv[1])
        sys.exit(1)

    # Check PS Script in working directory
    file_path_ps = os.path.join(os.getcwd(), PS_SCRIPT)
    if not os.path.isfile(file_path_ps):
        print_err(ERR_FILE_NOT_EXIST, PS_SCRIPT)
        sys.exit(1)

    # Generate csv files
    gen_user_csv_files(sys.argv[1], sys.argv[2], sys.argv[3])

    # Call powershell script as a sub-process
    try:
        result = sp.check_output(["gen-ad-users.exe", "users_new.csv"])
    except sp.CalledProcessError as e:
        print(f"AD Script failed with return code {e.returncode}")