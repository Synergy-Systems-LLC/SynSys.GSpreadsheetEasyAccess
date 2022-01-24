using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    [Serializable]
    internal class InvalidSheetNameException : Exception
    {
        public InvalidSheetNameException() { }

        public InvalidSheetNameException(string message) : base(message) { }

        public InvalidSheetNameException(string message, Exception innerException) : base(message, innerException) { }

        protected InvalidSheetNameException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}