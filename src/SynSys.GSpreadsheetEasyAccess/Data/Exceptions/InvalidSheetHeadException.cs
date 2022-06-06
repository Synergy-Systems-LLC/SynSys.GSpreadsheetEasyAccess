using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Represents an exception thrown due to the absence of at least one required column.
    /// </summary>
    [Serializable]
    public class InvalidSheetHeadException : Exception
    {
        /// <summary>
        /// Initializes a new InvalidSheetHeadException instance.
        /// </summary>
        public InvalidSheetHeadException() { }

        /// <summary>
        /// Initializes a new InvalidSheetHeadException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public InvalidSheetHeadException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new InvalidSheetHeadException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public InvalidSheetHeadException(string message, Exception innerException) : base(message, innerException) { }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected InvalidSheetHeadException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// The sheet whose state caused the exception.
        /// </summary>
        public SheetModel Sheet { get; internal set; }

        /// <summary>
        /// Required column headers that were not found in the sheet.
        /// </summary>
        public IEnumerable<string> LostedHeaders { get; internal set; }
    }
}