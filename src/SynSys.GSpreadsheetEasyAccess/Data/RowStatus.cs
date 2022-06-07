namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Determines the state of the row.<br/>
    /// The GCPApplication understands how to work with a specific row (based on this status)
    /// when sending data to Google spreadsheet.
    /// </summary>
    public enum RowStatus
    {
        /// <summary>
        /// Original row from Google spreadsheet.<br/>
        /// Assigned when loading from Google spreadsheet.
        /// </summary>
        Original,
        /// <summary>
        /// The row will be changed.
        /// </summary>
        ToChange,
        /// <summary>
        /// The row will be added.
        /// </summary>
        ToAppend,
        /// <summary>
        /// The row will be deleted.
        /// </summary>
        ToDelete,
    }
}
