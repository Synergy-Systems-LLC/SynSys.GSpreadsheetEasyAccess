import os
current_dir = os.path.abspath(__file__)
root_dir = os.path.dirname(os.path.dirname(os.path.dirname(current_dir)))
library_dir = os.path.join(root_dir, r'src\SynSys.GSpreadsheetEasyAccess\bin\Release')
library_name = 'SynSys.GSpreadsheetEasyAccess.dll'
if not os.path.exists(library_dir):
    raise Exception('Library path not found. Check this path "{}"'.format(library_dir))
if not os.path.exists(os.path.join(library_dir, library_name)):
    raise Exception(
        'Library {} not found. Build project SynSys.GSpreadsheetEasyAccess and start script again.'
        .format(library_name)
    )

import sys
sys.path.append(library_dir)

import clr
clr.AddReference(library_name)
from SynSys.GSpreadsheetEasyAccess import Application, Authentication, Data
from utils import read_credentials, print_sheet, change_sheet


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
uri = (
    'https://docs.google.com/spreadsheets/d/'
    '12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0'
)

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
