from enum import Enum


class Cell(object):
    """Represents a cell of the row."""

    @property
    def Value(self):
        return str()

    @Value.setter
    def Value(self, value):
        # type: (str) -> None
        """If the row in which this cell is located has the Original status,
        then the row status will change to ToChange while the cell value changes.
        """
        pass

    @property
    def Title(self):
        """The name of the column in which the cell is located."""
        return str()

    @property
    def Host(self):
        """Link to the row in which this cell is located."""
        return Row()

    @Host.setter
    def Host(self, value):
        # type: (Row) -> None
        """Link to the row in which this cell is located."""
        pass


class RowStatus(Enum):
    """Determines the state of the row.

    The GCPApplication understands how to work with a specific row (based on this status)
    when sending data to Google spreadsheet.
    """

    Original = 1
    """Original row from Google spreadsheet.
    Assigned when loading from Google spreadsheet.
    """

    ToChange = 2
    """The row will be changed."""

    ToAppend = 3
    """The row will be added."""

    ToDelete = 4
    """The row will be deleted."""


class Row(object):
    """The type represents one row of one Google spreadsheet sheet."""

    @property
    def Number(self):
        """Number, not index!"""
        return int()

    @property
    def Status(self):
        """Current status."""
        return RowStatus

    @property
    def Cells(self):
        """All cells in this row."""
        return [Cell()]

    @property
    def Key(self):
        """Key cell."""
        return Cell()


class SheetMode(Enum):
    """An enumeration for a specific sheet filling.

    The choice of mode affects the state of the cells and rows of the sheet.
    """

    Simple = 1
    """Spreadsheet without head and key."""

    Head = 2
    """Spreadsheet with head in one row."""

    HeadAndKey = 3
    """Spreadsheet with head in one row and and key column."""


class SheetModel(object):
    """The type represents one Google spreadsheet sheet."""

    @property
    def Title(self):
        """Sheet name."""
        return str()

    @property
    def SpreadsheetId(self):
        """Google spreadsheet Id."""
        return str()

    @property
    def Gid(self):
        """Google spreadsheet sheet Id."""
        return int()

    @property
    def SpreadsheetTitle(self):
        """Spreadsheet name."""
        return str()

    @property
    def KeyName(self):
        """Key column name."""
        return str()

    @property
    def Mode(self):
        """The mode by which work with the sheet is determined."""
        return SheetMode

    @property
    def Head(self):
        """First row of the sheet."""
        return [str()]

    @property
    def Rows(self):
        """All rows included in this sheet except for the head."""
        return [Row()]

    @property
    def IsEmpty(self):
        """Indicates that there are no rows in the sheet.

        If the sheet has a head, then it is not taken into account.
        """
        return bool()

    def AddRow(self, *args):
        """ Adds row to the end of the sheet.

        The row size will be equal to the maximum for this sheet and if the data is larger than
        this size, then part of the data will not be included into the sheet.\n
        If there is less data, then the remaining cells will be filled with empty values.

        Args:
            No args for add empty row.
        or Args:
            data (list[str]): Data to compose a row.
        """
        pass

    def DeleteRow(self, row):
        # type: (Row) -> None
        """The method deletes the row only if it had a RowStatus.ToAppend status.
        Otherwise, the method does not delete the row, but assigns the RowStatus.ToDelete status.\n
        This status will be taken into account when deleting rows from a Google spreadsheet sheet.

        Args:
            row (Row): Row to delete.
        """
        pass

    def Clean(self):
        """ Assigning a deletion status to all rows
        and physical deletion of rows that have not yet been added.
        """
        pass

    def CheckHead(self, requiredHeaders):
        # type: (list[str]) -> None
        """Check if the required columns exist in the spreadsheet.

        This method will throw an exception if the spreadsheet does not contain at least
        one required column.\n
        The method does not check for sheets with SheetMode.Simple.

        Args:
            requiredHeaders (list[str]): List of required sheet headers.

        Raises:
            InvalidSheetHeadException
        """
        pass

    def Merge(self, other):
        # type: (SheetModel) -> None
        """Merge with another version of the same sheet.

        Sheet is considered the same if it has the same basic characteristics
        except for the list of rows.\n
        Row comparison is performed before merging. Row changes occur after comparison if needed.\n
        Cell values and statuses are changed for rows, missing rows are added.

        Args:
            other (SheetModel): Same SheetModel.

        Raises:
            ArgumentNullException
            ArgumentException: Raise if other sheet not same.
        """
        pass


class Formatting(object):
    pass


class JsonSerialization(object):
    """Provides methods for converting a SheetModel instance
    to and from a JSON string representation.
    """

    @staticmethod
    def SerializeSheet(sheet, formatting):
        # type: (SheetModel, Formatting) -> str
        """Serializes an instance of the specified sheet to JSON using the selected formatting.

        Returns:
            The string representation of the object in the format JSON.
        """
        return str()

    @staticmethod
    def DeserializeSheet(jsonSheet):
        # type: (str) -> SheetModel
        """Deserializes JSON to an instance SheetModel.

        Returns:
            Deserialized object from string JSON.
        """
        return SheetModel()
