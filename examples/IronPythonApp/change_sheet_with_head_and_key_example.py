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
from SynSys.GSpreadsheetEasyAccess.Application.Exceptions import (
    UserAccessDeniedException, SpreadsheetNotFoundException, SheetNotFoundException
)
from SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions import (
    AuthenticationTimedOutException, UserCanceledAuthenticationException
)
from SynSys.GSpreadsheetEasyAccess.Data.Exceptions import (
    SheetKeyNotFoundException, InvalidSheetHeadException, EmptySheetException
)
from utils import read_credentials, print_sheet, change_sheet


try:
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
    # Don't forget change change_sheet function.
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
except AuthenticationTimedOutException as e:
    print(
        "Authentication is timed out.\n"
        "Run the plugin again and authenticate in the browser."
    )
except UserCanceledAuthenticationException as e:
    print(
        "You have canceled authentication.\n"
        "You need to be authenticated to use library SynSys.GSpreadsheetEasyAccess."
    )
except UserAccessDeniedException as e:
    print(
        "User is denied access to the operation: {}\n"
        "Reason: {}\n"
        "Check table access.\n"
        "The table must be available to all users who have the link.".format(e.Operation, e.Message)
    )
except SpreadsheetNotFoundException as e:
    print(
        "Failed to get spreadsheet with\n"
        "spreadsheet id: {}\n"
        "Check if this table exists.".format(e.SpreadsheetId)
    )
except SheetNotFoundException as e:
    print(
        "Failed to get sheet with\n"
        "spreadsheet id: {}\n"
        "spreadsheet title: {}\n"
        "sheet id: {}\n"
        "Check if this sheet exists.".format(e.SpreadsheetId, e.SpreadsheetName, e.SheetGid)
    )
except SheetKeyNotFoundException as e:
    print(
        "Key column \"{}\" require in the sheet\n"
        "spreadsheet: \"{}\"\n"
        "sheet: \"{}\"\n"
        "for correct plugin operation\n".format(
            e.Sheet.KeyName, e.Sheet.SpreadsheetTitle, e.Sheet.Title
        )
    )
except InvalidSheetHeadException as e:
    print(
        "Spreadsheet \"{}\"\n"
        "sheet \"{}\"\n"
        "lacks required headers:\n"
        "{}.".format(e.Sheet.SpreadsheetTitle, e.Sheet.Title, ";\n".join(e.LostHeaders))
    )
except EmptySheetException as e:
    print(
        "For the plugin to work correctly sheet"
        "spreadsheet: \"{}\"\n"
        "sheet: \"{}\"\n"
        "cannot be empty.".format(e.Sheet.SpreadsheetTitle, e.Sheet.Title)
    )
except Exception as e:
    print "An unhandled exception occurred.\n{}".format(e)  # type: ignore
