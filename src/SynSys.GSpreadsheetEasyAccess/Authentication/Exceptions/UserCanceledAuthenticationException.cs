using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions
{
    /// <summary>
    /// Represents an exception thrown because the user canceled the authentication process.
    /// </summary>
    [Serializable]
    public class UserCanceledAuthenticationException : Exception
    {
        /// <summary>
        /// Initializes a new UserCanceledAuthenticationException instance.
        /// </summary>
        public UserCanceledAuthenticationException() { }

        /// <summary>
        /// Initializes a new UserCanceledAuthenticationException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public UserCanceledAuthenticationException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new UserCanceledAuthenticationException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public UserCanceledAuthenticationException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected UserCanceledAuthenticationException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}