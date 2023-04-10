from SynSys.GSpreadsheetEasyAccess.Data import SheetModel


class EmptySheetException(Exception):
    """Represents an exception thrown due to invalid data in a Google spreadsheet sheet."""

    @property
    def Sheet(self):
        """The sheet whose state caused the exception."""
        return SheetModel()

    @Sheet.setter
    def Sheet(self, value):
        """The sheet whose state caused the exception."""
        pass


class InvalidSheetHeadException(Exception):
    """Represents an exception thrown due to the absence of at least one required column."""

    @property
    def Sheet(self):
        """The sheet whose state caused the exception."""
        return SheetModel()

    @property
    def LostedHeaders(self):
        """Required column headers that were not found in the sheet."""
        return [str()]


class SheetKeyNotFoundException(Exception):
    """Represents an exception thrown due to missing key column."""

    @property
    def Sheet(self):
        """The sheet whose state caused the exception."""
        return SheetModel()

    @Sheet.setter
    def Sheet(self, value):
        """The sheet whose state caused the exception."""
        pass
