import gspread
from google.oauth2.service_account import Credentials
import google.auth
import re
import string

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

class SheetsEngine:
    """
    Handles everything directly related to the API
    """
    credentials = Credentials.from_service_account_file('credentials.json', scopes=scope)
    gc = gspread.authorize(credentials)

    @staticmethod
    def get_service_email():
        return SheetsEngine.credentials.service_account_email
    
    @staticmethod
    def get_values_from_key(key):
        return SheetsEngine.gc.open_by_key(key).sheet1.get_all_values()

    @staticmethod
    def append_to_sheet(key, rows):
        SheetsEngine.gc.open_by_key(key).sheet1.append_rows(rows)



def get_sheets_keys():
    """
    Gets the sheet keys and the column to merge on from the user
    input formats:
    num_of_sheets: Integer > 1
    sheet: string of alphanumeric characters and underscores only
    col: Letters only or integer > 0
    """
    num_of_sheets = None
    sheets = []
    # Get the number of sheets to merge
    while True:
        try:
            num_of_sheets = int(input("Number of sheets to merge: "))
            assert num_of_sheets > 1, "You can't merge less than 2 sheets"
            break
        except ValueError:
            print("Invalid Input (Must be integer and greater than 1)")
            continue
    # Get the sheet keys
    print("*Sheet keys can be found in the url of the spreadsheet")
    for i in range(num_of_sheets):
        while True:
            key = input(f'Enter Sheet#{i+1} key: ')
            if re.match("^[A-Za-z0-9_-]*$", key):
                sheets.append(key)
                break
            else:
                print("Invalid key format")
                continue
    # Get the column ID, it can be letters only or a number
    while True:
        col = input('Enter column to merge on: ')
        if (re.match("^[A-Za-z]*$", col) or re.match("^[0-9]*$", col)) and col[0] != '0' :
            break
        else:
            print("Invalid Column ID (Either letters only or integer > 0)")
            continue
    return [sheets, col_to_num(col)]


def col_to_num(col):
    """
    Converts the col from letters to an integer
    """
    if re.match("^[A-Za-z]*$", col):
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num-1
    else:
        return int(col)-1

def keys_to_sheets(keys):
    """
    For every key fetch the corresponding sheet values as a list of rows
    and if it couldn't fetch a single sheet it aborts
    """
    sheets = []
    for i, key in enumerate(keys):
        try:
            sheet_values = SheetsEngine.get_values_from_key(key)
            sheets.append(sheet_values)
        except gspread.exceptions.APIError:
            print(f"Couldn't fetch sheet #{i+1} whose key is {key},\
             please make sure the key is valid\
              and the document is shared with {SheetsEngine.get_service_email()} and try again")
            exit(1)
    return sheets


def merge_sheets(main_key, main_sheet, sheets, col):
    """
    Takes a main sheet and multiple other sheets to merge them on a certain col
    """
    unique_column_values = set([row[col] for row in main_sheet])
    new_rows = []
    for sheet in sheets:
        for row in sheet:
            if row[col] not in unique_column_values:
                new_rows.append(row)
                unique_column_values.add(row[col])
    try:
        SheetsEngine.append_to_sheet(main_key, new_rows)
        print('Merged successfully')
    except:
        print('Merging failed while appending to the first sheet')



def main():
    sheets_keys, col = get_sheets_keys()
    sheets_values = keys_to_sheets(sheets_keys)
    main_key, main_values = sheets_keys[0], sheets_values[0]
    merge_sheets(main_key, main_values, sheets_values[1:], col)

    

if __name__ == "__main__":
    main()