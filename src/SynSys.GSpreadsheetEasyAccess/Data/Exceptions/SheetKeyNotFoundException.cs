using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Represents an exception thrown due to missing key column.
    /// </summary>
    [Serializable]
    public class SheetKeyNotFoundException : Exception
    {
        /// <summary>
        /// Initializes a new SheetKeyNotFoundException instance.
        /// </summary>
        public SheetKeyNotFoundException() { }

        /// <summary>
        /// Initializes a new SheetKeyNotFoundException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public SheetKeyNotFoundException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new SheetKeyNotFoundException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SheetKeyNotFoundException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected SheetKeyNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// The sheet whose state caused the exception.
        /// </summary>
        public UserSheet Sheet { get; set; }
    }
}