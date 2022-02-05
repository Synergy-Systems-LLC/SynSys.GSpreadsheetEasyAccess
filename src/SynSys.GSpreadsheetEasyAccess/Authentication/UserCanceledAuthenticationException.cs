using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    [Serializable]
    internal class UserCanceledAuthenticationException : Exception
    {
        public UserCanceledAuthenticationException() { }

        public UserCanceledAuthenticationException(string message) : base(message) { }


        public UserCanceledAuthenticationException(string message, Exception innerException) : base(message, innerException) { }

        protected UserCanceledAuthenticationException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}