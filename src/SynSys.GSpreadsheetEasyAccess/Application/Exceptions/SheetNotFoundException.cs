using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    [Serializable]
    internal class SheetNotFoundException : Exception
    {
        public SheetNotFoundException() { }

        public SheetNotFoundException(string message) : base(message) { }

        public SheetNotFoundException(string message, Exception innerException) : base(message, innerException) { }

        protected SheetNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}