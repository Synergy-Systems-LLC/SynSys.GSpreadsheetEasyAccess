using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Represents an exception that occurred due to the inability to create a sheet.
    /// </summary>
    [Serializable]
    public class CreatingSheetException : Exception
    {
        /// <summary>
        /// Initializes a new CreatingSheetException instance.
        /// </summary>
        public CreatingSheetException() { }

        /// <summary>
        /// Initializes a new CreatingSheetException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public CreatingSheetException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new CreatingSheetException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public CreatingSheetException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected CreatingSheetException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// Id of the spreadsheet in which the sheet could not be created
        /// </summary>
        public string SpreadsheetId { get; internal set; }

        /// <summary>
        /// Title of the spreadsheet in which the sheet could not be created
        /// </summary>
        public string SpreadsheetTitle { get; internal set; }

        /// <summary>
        /// Title of the sheet that could not be created in the spreadsheet.
        /// </summary>
        public string SheetTitle { get; internal set; }
    }
}