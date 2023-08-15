namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Determines the state of google sheet members.<br/>
    /// The GCPApplication understands how to work with a specific member (based on this status)
    /// when sending data to Google spreadsheet or working with model of sheet.
    /// </summary>
    public enum ChangeStatus
    {
        /// <summary>
        /// Original row from Google spreadsheet.<br/>
        /// Assigned when loading from Google spreadsheet.
        /// </summary>
        Original,
        /// <summary>
        /// Will be changed.
        /// </summary>
        ToChange,
        /// <summary>
        /// Will be added.
        /// </summary>
        ToAppend,
        /// <summary>
        /// Will be inserted.
        /// </summary>
        ToInsert,
        /// <summary>
        /// Will be deleted.
        /// </summary>
        ToDelete,
    }
}
