using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions
{
    /// <summary>
    /// Represents an exception thrown due to authentication timeout.
    /// </summary>
    [Serializable]
    public class AuthenticationTimedOutException : Exception
    {
        /// <summary>
        /// Initializes a new AuthenticationTimedOutException instance.
        /// </summary>
        public AuthenticationTimedOutException() { }

        /// <summary>
        /// Initializes a new AuthenticationTimedOutException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public AuthenticationTimedOutException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new AuthenticationTimedOutException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public AuthenticationTimedOutException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected AuthenticationTimedOutException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}