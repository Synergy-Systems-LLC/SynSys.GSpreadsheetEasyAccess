import os
current_dir = os.path.abspath(__file__)
root_dir = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))
library_dir = os.path.join(root_dir, r'src\SynSys.GSpreadsheetEasyAccess\bin\Release')
if not os.path.exists(library_dir):
    raise Exception('Library path not found. Check this path "{}"'.format(library_dir))

import sys
sys.path.append(library_dir)

import clr
import System
from System.Collections.Generic import List

clr.AddReference('SynSys.GSpreadsheetEasyAccess.dll')
from SynSys.GSpreadsheetEasyAccess import Application, Authentication, Data


def read_credentials(path):
    # type: (str) -> List[System.Byte]
    """Return the contents of a file"""
    with open(path, 'rb') as file:
        return System.Text.Encoding.ASCII.GetBytes(''.join(file.readlines()))


def print_sheet(sheet, status):
    # type: (str, str) -> None
    print(
        "\n"
        "Status:           {}\n"
        "Spreadsheet Name: {}\n"
        "Sheet Name:       {}\n"
        "Number of lines:  {}\n".format(
            status, sheet.SpreadsheetTitle, sheet.Title, sheet.Rows.Count
        )
    )
    print ' {:>3}'.format(''),  # type: ignore
    for title in sheet.Head:
        if not title:
            title = '-'
        print '|{:>7}'.format(title),  # type: ignore
    print '| '  # type: ignore
    print ' {:>3}'.format(''),  # type: ignore
    print '|{}'.format('-' * 8) * len(sheet.Head) + '|'  # type: ignore
    for row in sheet.Rows:
        print '|{:>3}'.format(row.Number),  # type: ignore
        for cell in row.Cells:
            print '|{:>7}'.format(cell.Value),  # type: ignore
        print '| {}'.format(row.Status)  # type: ignore


def change_sheet(sheet):
    add_rows(sheet)
    change_rows(sheet)
    delete_rows(sheet)


def add_rows(sheet):
    # Add empty row.
    sheet.AddRow()

    # Add a row part of which will be filled with empty cells.
    sheet.AddRow(List[str](["123", "asd"]))

    # Add a row where part of the data will not fall into the line, namely "k" and "l",
    # because test sheet has 10 columns.
    sheet.AddRow(List[str](["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l"]))


def change_rows(sheet):
    # This example doesn't take into account the absence of a key with the desired value
    # and a cell with the selected Title
    sheet.Rows.Find(lambda row: row.Key.Value == "31") \
        .Cells.Find(lambda cell: cell.Title == "Head 6") \
        .Value = "360"
    sheet.Rows.Find(lambda row: row.Key.Value == "61") \
        .Cells.Find(lambda cell: cell.Title == "Head 3") \
        .Value = "630"
    sheet.Rows.Find(lambda row: row.Key.Value == "51") \
        .Cells.Find(lambda cell: cell.Title == "Head 2") \
        .Value = "520"

    # Change added rows with status RowStatus.ToAppend
    list(sheet.Rows)[-1].Cells[0].Value = "change"
    list(sheet.Rows)[-1].Cells[1].Value = "added"
    list(sheet.Rows)[-1].Cells[2].Value = "line"


def delete_rows(sheet):
    # This example doesn't take into account the lack of necessary indexes
    sheet.DeleteRow(sheet.Rows[3])

    # Delete the added line with the status RowStatus.ToAppend.
    # In this case, the line is immediately deleted without waiting for the sheet to be updated.
    sheet.DeleteRow(sheet.Rows[10])


# Initialization of the class representing the application on the Google Cloud Platform
gcp_app = Application.GCPApplication()

# To change data in Google spreadsheets,
# you need to authenticate a real user and designate access rights.
gcp_app.AuthenticateAs(
    Authentication.UserAccount(
        read_credentials(r'examples\IronPythonApp\credentials.json'),
        Authentication.OAuthSheetsScope.FullAccess
    )
)

# For tests, you need to change the given uri to your own.
# Don't forget change change_cheet function.
uri = 'https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0'

sheet = gcp_app.GetSheetWithHeadAndKey(uri, 'Head 1')  # type: Data.SheetModel
print_sheet(sheet, 'Original sheet')

change_sheet(sheet)
print_sheet(sheet, 'Changed sheet before updating to google')

# In order for the data in the google spreadsheet sheet to be updated,
# you need to pass the changed instance of the SheetModel to the UpdateSheet method
gcp_app.UpdateSheet(sheet)
print_sheet(sheet, 'Sheet after update in google')

# The function is called a second time to show
# that the current instance of the sheet can still be used after the update.
# Its structure will match the google spreadsheet.
change_sheet(sheet)
print_sheet(sheet, 'Changed sheet before updating to google')

gcp_app.UpdateSheet(sheet)
print_sheet(sheet, 'Sheet after update in google')
