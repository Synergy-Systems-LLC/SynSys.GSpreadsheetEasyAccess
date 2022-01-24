using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    [Serializable]
    internal class EmptySheetException : Exception
    {
        public EmptySheetException() { }

        public EmptySheetException(string message) : base(message) { }

        public EmptySheetException(string message, Exception innerException) : base(message, innerException) { }

        protected EmptySheetException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}