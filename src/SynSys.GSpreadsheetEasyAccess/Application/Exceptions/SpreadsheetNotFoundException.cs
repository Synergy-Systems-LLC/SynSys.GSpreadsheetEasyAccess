using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    [Serializable]
    internal class SpreadsheetNotFoundException : Exception
    {
        public SpreadsheetNotFoundException() { }

        public SpreadsheetNotFoundException(string message) : base(message) { }

        public SpreadsheetNotFoundException(string message, Exception innerException) : base(message, innerException) { }

        protected SpreadsheetNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}