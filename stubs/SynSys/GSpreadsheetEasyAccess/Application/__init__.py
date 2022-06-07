from SynSys.GSpreadsheetEasyAccess.Authentication import Principal
from SynSys.GSpreadsheetEasyAccess.Data import SheetModel


class GCPApplication(object):
    """Represents an application on the Google Cloud Platform that has access to \
    [Google Sheets API](https://developers.google.com/sheets/api?hl=en_US).

    The class serves to receive and update data from Google Sheets.\n
    Methods can only be used after successful authentication.
    """

    def AuthenticateAs(self, principal):
        # type: (Principal) -> None
        """ To gain access to the Google Sheets API, you must be authenticated.
        It is necessary to specify who is authenticating.
        """
        pass

    def GetSheet(uri):
        # type: (str) -> SheetModel
        """ Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.

        Args:
            uri (str): Full uri of spreadsheet sheet.
        or Args:
            spreadsheetId (str): Spreadhseet Id.
            gid (int): Spreadsheet sheet Id.
        or Args:
            spreadsheetId (str): Spreadhseet Id.
            sheetName (str): Spreadsheet sheet name.

        Returns:
            SheetModel is a list of Rows.\n
            Header is absent.\n
            Each row has the same number of cells.\n
            Each cell has a string value.
        """
        return SheetModel()

    def GetSheetWithHead(self, uri):
        # type: (str) -> SheetModel
        """ Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.

        Args:
            uri (str): Full uri of spreadsheet sheet.
        or Args:
            spreadsheetId (str): Spreadhseet Id.
            gid (int): Spreadsheet sheet Id.
        or Args:
            spreadsheetId (str): Spreadhseet Id.
            sheetName (str): Spreadsheet sheet name.

        Returns:
            SheetModel is a list of Rows without first row.\n
            First row is a header of sheet.\n
            Each row has the same number of cells.\n
            Each cell has a string value and title,
            which matches the column heading for the given cell.
        """
        return SheetModel()

    def GetSheetWithHeadAndKey(self, uri, keyName):
        # type: (str, str) -> SheetModel
        """ Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.

        Args:
            uri (str): Full uri of spreadsheet sheet.
            keyName (str): sheet key column.
        or Args:
            spreadsheetId (str): Spreadhseet Id.
            gid (int): Spreadsheet sheet Id.
            keyName (str): sheet key column.
        or Args:
            spreadsheetId (str): Spreadhseet Id.
            sheetName (str): Spreadsheet sheet name.
            keyName (str): sheet key column.

        Returns:
            SheetModel is a list of Rows without first row.\n
            First row is a header of sheet.\n
            Each row has the same number of cells and has key column.\n
            Each cell has a string value and title,
            which matches the column heading for the given cell.
        """
        return SheetModel()

    def UpdateSheet(self, sheet):
        # type: (SheetModel) -> None
        """Update the Google spreadsheet sheet based on the modified instance of the SheetModel type.

        The method changes the data in the cells,
        adds rows to the end of the sheet and removes the selected rows.\n
        All these actions are based on requests to Google.

        Args:
            sheetModel (SheetModel): Google spreadsheet sheet model.

        Returns:
            SheetModel: Updated instance of SheetModel.
        """
        return SheetModel()


class HttpUtils(object):
    """Provide static methods for parsing Google spreadsheets uri."""

    @staticmethod
    def GetSpreadsheetIdFromUri(uri):
        # type: (str) -> str
        """Get the Google spreadsheet Id.

        Args:
            uri (str): Full Google spreadsheet sheet uri.

        Returns:
            str: String of 44 characters following d/.
        """
        return str()

    @staticmethod
    def GetGidFromUri(uri):
        # type: (str) -> int
        """Get the Google spreadsheet sheet Id.

        Args:
            uri (str): Full Google spreadsheet sheet uri.

        Returns:
            int: Returns -1 if there is no gid in uri.
        """
        return int()
