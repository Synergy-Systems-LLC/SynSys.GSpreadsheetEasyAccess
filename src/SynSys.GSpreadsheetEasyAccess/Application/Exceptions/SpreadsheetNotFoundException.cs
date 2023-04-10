using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Represents the exception thrown due to table could not be found by the current spreadsheet Id.
    /// </summary>
    [Serializable]
    public class SpreadsheetNotFoundException : Exception
    {
        /// <summary>
        /// Initializes a new SpreadsheetNotFoundException instance.
        /// </summary>
        public SpreadsheetNotFoundException() { }

        /// <summary>
        /// Initializes a new SpreadsheetNotFoundException instance with an message about exception.
        /// </summary>
        /// <param name="message"></param>
        public SpreadsheetNotFoundException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new SpreadsheetNotFoundException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SpreadsheetNotFoundException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected SpreadsheetNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// Spreadsheet id that caused the exception.
        /// </summary>
        public string SpreadsheetId { get; internal set; }
    }
}