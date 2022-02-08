using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    [Serializable]
    internal class UserAccessDeniedException : Exception
    {
        public UserAccessDeniedException() { }

        public UserAccessDeniedException(string message) : base(message) { }

        public UserAccessDeniedException(string message, Exception innerException) : base(message, innerException) { }

        protected UserAccessDeniedException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}