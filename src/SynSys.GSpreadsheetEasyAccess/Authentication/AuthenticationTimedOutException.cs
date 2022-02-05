using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    [Serializable]
    internal class AuthenticationTimedOutException : Exception
    {
        public AuthenticationTimedOutException() { }

        public AuthenticationTimedOutException(string message) : base(message) { }


        public AuthenticationTimedOutException(string message, Exception innerException) : base(message, innerException) { }

        protected AuthenticationTimedOutException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}