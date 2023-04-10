using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Represents the exception thrownbecause the sheet could not be found by the current sheet Id.
    /// </summary>
    [Serializable]
    public class SheetNotFoundException : Exception
    {
        /// <summary>
        /// Initializes a new SheetNotFoundException instance.
        /// </summary>
        public SheetNotFoundException() { }

        /// <summary>
        /// Initializes a new SheetNotFoundException instance with an message about exception.
        /// </summary>
        /// <param name="message"></param>
        public SheetNotFoundException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new SheetNotFoundException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SheetNotFoundException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected SheetNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// The name of the spreadsheet in which the sheet was searched.
        /// </summary>
        public string SpreadsheetTitle { get; internal set; }

        /// <summary>
        /// Id of the Spreadsheet in which the sheet was searched.
        /// </summary>
        public string SpreadsheetId { get; internal set; }

        /// <summary>
        /// Id of the sheet by which the sheet could not be found.
        /// </summary>
        public string SheetGid { get; internal set; }

        /// <summary>
        /// Title of the sheet by which the sheet could not be found.
        /// </summary>
        public string SheetTitle { get; internal set; }
    }
}