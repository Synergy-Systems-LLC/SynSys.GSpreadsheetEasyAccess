using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    [Serializable]
    internal class InvalidSheetUriException : Exception
    {
        public InvalidSheetUriException() { }

        public InvalidSheetUriException(string message) : base(message) { }

        public InvalidSheetUriException(string message, Exception innerException) : base(message, innerException) { }

        protected InvalidSheetUriException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}