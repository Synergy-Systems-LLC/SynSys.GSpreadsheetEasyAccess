using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess
{
    [Serializable]
    internal class InvalidSheetGidException : Exception
    {
        public InvalidSheetGidException() { }

        public InvalidSheetGidException(string message) : base(message) { }

        public InvalidSheetGidException(string message, Exception innerException) : base(message, innerException) { }

        protected InvalidSheetGidException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}