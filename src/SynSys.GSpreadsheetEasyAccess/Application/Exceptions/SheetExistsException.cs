using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Represents an exception thrown due to the existence of a sheet with that title in this Google spreadsheet.
    /// </summary>
    [Serializable]
    public class SheetExistsException : Exception
    {
        /// <summary>
        /// Initializes a new SheetExistsException instance.
        /// </summary>
        public SheetExistsException() { }

        /// <summary>
        /// Initializes a new SheetExistsException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public SheetExistsException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new SheetExistsException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SheetExistsException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected SheetExistsException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// Id of the spreadsheet where the sheet exists.
        /// </summary>
        public string SpreadsheetId { get; set; }

        /// <summary>
        /// Title of the spreadsheet where the sheet exists.
        /// </summary>
        public string SpreadsheetTitle { get; set; }

        /// <summary>
        /// Title of the sheet that already exists in the Google table
        /// </summary>
        public string SheetTitle { get; set; }
    }
}