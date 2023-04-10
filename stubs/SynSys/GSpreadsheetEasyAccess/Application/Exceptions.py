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
    def SpreadsheetTitle(self):
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
    def SheetTitle(self):
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
    """Represents an exception thrown due to given user has been denied access
    to any actions with Google spreadsheet.
    """

    @property
    def Operation(self):
        """Contains the name of the operation to which the user does not have access."""
        return str()

    @property
    def Message(self):
        return str()


class SheetExistsException(Exception):
    """Represents an exception thrown due to the existence of a
    sheet with that title in this Google spreadsheet."""

    @property
    def SpreadsheetId(self):
        """Id of the spreadsheet where the sheet exists."""
        return str()

    @property
    def SpreadsheetTitle(self):
        """Title of the spreadsheet where the sheet exists."""
        return str()

    @property
    def SheetTitle(self):
        """Title of the sheet that already exists in the Google table."""
        return str()


class CreatingSheetException(Exception):
    """Represents an exception that occurred due to the inability to create a sheet."""

    @property
    def SpreadsheetId(self):
        """Id of the spreadsheet where the sheet exists."""
        return str()

    @property
    def SpreadsheetTitle(self):
        """Title of the spreadsheet where the sheet exists."""
        return str()

    @property
    def SheetTitle(self):
        """Title of the sheet that already exists in the Google table."""
        return str()
