using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application
{
    [Serializable]
    internal class InvalidApiKeyException : Exception
    {
        public InvalidApiKeyException() { }

        public InvalidApiKeyException(string message) : base(message) { }

        public InvalidApiKeyException(string message, Exception innerException) : base(message, innerException) { }

        protected InvalidApiKeyException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}