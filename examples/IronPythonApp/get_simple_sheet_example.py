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
from utils import read_apikey, print_sheet


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
uri = (
    'https://docs.google.com/spreadsheets/d/'
    '12nBUl0szLwBJKfWbe6aA1bNGxNwUJzNwvXyRSPkS8io/edit#gid=0'
)

sheet = gcp_app.GetSheet(uri)  # type: Data.SheetModel
print_sheet(sheet, 'Original sheet')
