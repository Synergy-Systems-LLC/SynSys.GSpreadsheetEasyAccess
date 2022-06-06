using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Represents an exception thrown due to invalid data in a Google spreadsheet sheet.
    /// </summary>
    [Serializable]
    public class EmptySheetException : Exception
    {
        /// <summary>
        /// Initializes a new EmptySheetException instance.
        /// </summary>
        public EmptySheetException() { }

        /// <summary>
        /// Initializes a new EmptySheetException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public EmptySheetException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new EmptySheetException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public EmptySheetException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected EmptySheetException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// The sheet whose state caused the exception
        /// </summary>
        public SheetModel Sheet { get; set; }
    }
}