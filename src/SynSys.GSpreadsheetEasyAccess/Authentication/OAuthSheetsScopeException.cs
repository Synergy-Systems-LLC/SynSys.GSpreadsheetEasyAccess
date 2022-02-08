using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    [Serializable]
    internal class OAuthSheetsScopeException : Exception
    {
        public OAuthSheetsScopeException() { }

        public OAuthSheetsScopeException(string message) : base(message) { }

        public OAuthSheetsScopeException(string message, Exception innerException) : base(message, innerException) { }

        protected OAuthSheetsScopeException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
