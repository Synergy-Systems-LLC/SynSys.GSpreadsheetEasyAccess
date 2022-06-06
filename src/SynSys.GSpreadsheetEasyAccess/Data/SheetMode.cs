namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// An enumeration for a specific sheet filling.
    /// The choice of mode affects the state of the cells and rows of the sheet.
    /// </summary>
    public enum SheetMode
    {
        /// <summary>
        /// Spreadsheet without head and key.
        /// </summary>
        Simple,
        /// <summary>
        /// Spreadsheet with head in one row.
        /// </summary>
        Head,
        /// <summary>
        /// Spreadsheet with head in one row and and key column.
        /// </summary>
        HeadAndKey
    }
}
