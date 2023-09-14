import System
from SynSys.GSpreadsheetEasyAccess.Data import SheetModel


def read_api_key(path):
    # type: (str) -> str
    """Return the contents of a file"""
    with open(path, 'r') as file:
        return ''.join(file.readlines())


def read_credentials(path):
    # type: (str) -> list[System.Byte]
    """Return the contents of a file"""
    with open(path, 'rb') as file:
        return System.Text.Encoding.ASCII.GetBytes(''.join(file.readlines()))  # type: ignore


def print_sheet(sheet, status):
    # type: (SheetModel, str) -> None
    print(
        "\n"
        "Status:            {}\n"
        "Spreadsheet Title: {}\n"
        "Sheet Title:       {}\n"
        "Number of lines:   {}\n".format(
            status, sheet.SpreadsheetTitle, sheet.Title, sheet.Rows.Count  # type: ignore
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
    sheet.AddRow(["123", "asd"])

    # Add a row where part of the data will not fall into the line, namely "k" and "l",
    # because test sheet has 10 columns.
    sheet.AddRow(["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l"])


def change_rows(sheet):
    # This example doesn't take into account the absence of a key with the desired value
    # and a cell with the selected Title
    sheet.Rows.Find(lambda row: row.Key.Value == "31") \
        .Cells.Find(lambda cell: cell.Title == "Title 6") \
        .Value = "360"
    sheet.Rows.Find(lambda row: row.Key.Value == "61") \
        .Cells.Find(lambda cell: cell.Title == "Title 3") \
        .Value = "630"
    sheet.Rows.Find(lambda row: row.Key.Value == "51") \
        .Cells.Find(lambda cell: cell.Title == "Title 2") \
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
