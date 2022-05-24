import os
current_dir = os.path.abspath(__file__)
root_dir = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))
library_dir = os.path.join(root_dir, r'src\SynSys.GSpreadsheetEasyAccess\bin\Release')
if not os.path.exists(library_dir):
    raise Exception('Library path not found. Check this path "{}"'.format(library_dir))

import sys
sys.path.append(library_dir)

import clr
clr.AddReference('SynSys.GSpreadsheetEasyAccess.dll')
from SynSys.GSpreadsheetEasyAccess import Application, Authentication, Data


def read_apikey(path):
    # type: (str) -> str
    """Return the contents of a file"""
    with open(path, 'r') as file:
        return ''.join(file.readlines())


def print_sheet(sheet, status):
    # type: (str, str) -> None
    print(
        "\n"
        "Status:           {}\n"
        "Spreadsheet Name: {}\n"
        "Sheet Name:       {}\n"
        "Number of lines:  {}".format(status, sheet.SpreadsheetTitle, sheet.Title, sheet.Rows.Count)
    )
    for row in sheet.Rows:
        print '|{:>3}'.format(row.Number),  # type: ignore
        for cell in row.Cells:
            print '|{:>7}'.format(cell.Value),  # type: ignore
        print '| {}'.format(row.Status)  # type: ignore


# Initialization of the class representing the application on the Google Cloud Platform.
gcp_app = Application.GCPApplication()

# To use the google spreadsheet service, you need to authenticate the user.
# To get data, a service account is enough, but the spreadsheet must be open to everyone.
gcp_app.AuthenticateAs(
    Authentication.ServiceAccount(
        read_apikey(r'examples\IronPythonApp\apikey.txt')
    )
)

# This uri can be used by everyone because the table is open to everyone.
uri = 'https://docs.google.com/spreadsheets/d/12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0'

sheet = gcp_app.GetSheet(uri)  # type: Data.SheetModel
print_sheet(sheet, 'Original sheet')
