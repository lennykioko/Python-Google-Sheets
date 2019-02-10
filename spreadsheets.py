"""
Integrates with google spreadsheets.
You may need to install PyOpenSSL if app does not work on your end.
The rows and cols  are not zero-index based, they read as on the spreadsheet.
"""
from pprint import pprint as pp

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Authorization

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

credentials = ServiceAccountCredentials.from_json_keyfile_name(
    "python-sheets-docs-drive-a29d35cbc82b.json", scope)

gs = gspread.authorize(credentials)


# Working with your spreadsheets

# open the spreadsheet to work with
worksheet = gs.open("TEST PYTHON INTEGRATION").sheet1


def append_row(value_list):
    """Enter a list of the data you wish to populate the row with"""
    worksheet.append_row(value_list)


def delete_row(row_num):
    """Enter the index of the row to be deleted"""
    worksheet.delete_row(row_num)


def get_row_by_acell(acell):
    """Example of a valid acell value is A2"""
    cell_object = worksheet.acell(acell)
    return cell_object


def get_row_by_cell(row, col):
    """The row and col here are not zero-index based"""

    cell_object = worksheet.cell(row, col)
    return cell_object


def get_cell_value(cell_object):
    return cell_object.value


def get_cell_row(cell_object):
    return cell_object.row


def get_cell_col(cell_object):
    return cell_object.col


def update_acell(acell, new_value):
    updated_cell_object = worksheet.update_acell(acell, new_value)
    return updated_cell_object


def update_cell(row, col, new_value):
    updated_cell_object = worksheet.update_cell(row, col, new_value)
    return updated_cell_object


def find(search_term):
    """Ensure that the search_term is a string"""
    cell_object = worksheet.find(search_term)
    return cell_object


def find_all(search_term):
    """Ensure that the search_term is a string"""
    cell_objects = worksheet.findall(search_term)
    return cell_objects


def update_one_in_queryset(search_term, index_to_update, new_value):
    list_of_items = worksheet.findall(search_term)
    list_of_items[index_to_update].value = new_value

    # checks for any update to the queryset and updates the changed record
    worksheet.update_cells(list_of_items)

    return f"""Successfully updated {search_term}
            at index {index_to_update} to {new_value}"""


def update_all_in_queryset(search_term, new_value):
    list_of_items = worksheet.findall(search_term)
    for item in list_of_items:
        item.value = new_value

    # checks for any update to the queryset and updates the changed records
    worksheet.update_cells(list_of_items)

    return f"""Successfully updated all instances of {search_term}
            to {new_value}"""


def get_all_records():
    return worksheet.get_all_records()


def pp_records(records):
    """Pretty print records"""
    pp(records)
