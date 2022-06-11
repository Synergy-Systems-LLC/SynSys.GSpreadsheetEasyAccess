class InvalidApiKeyException(Exception):
    """Represents an exception thrown due to an invalid API key of the application
    on the Google Cloud Platform.
    """
    pass


class SheetNotFoundException(Exception):
    """Represents the exception thrown because the sheet could not be found
    by the current sheet Id.
    """

    @property
    def SpreadsheetName(self):
        """The name of the spreadsheet in which the sheet was searched."""
        return str()

    @property
    def SpreadsheetId(self):
        """Id of the Spreadsheet in which the sheet was searched."""
        return str()

    @property
    def SheetGid(self):
        """Id of the sheet by which the sheet could not be found."""
        return str()

    @property
    def SheetName(self):
        """The name of the sheet by which the sheet could not be found."""
        return str()


class SpreadsheetNotFoundException(Exception):
    """Represents the exception thrown because the table could not be found
    by the current spreadsheet id.
    """

    @property
    def SpreadsheetId(self):
        """Spreadsheet id that caused the exception."""
        return str()


class UserAccessDeniedException(Exception):
    """Represents an exception thrownbecause the given user has been denied access
    to any actions with Google spreadsheet.
    """

    @property
    def Operation(self):
        """Contains the name of the operation to which the user does not have access."""
        return str()
