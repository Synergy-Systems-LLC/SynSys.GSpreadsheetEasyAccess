using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions
{
    /// <summary>
    /// Represents an exception throwndue to a mismatch between the scopes of the application and the requests being sent
    /// </summary>
    [Serializable]
    public class OAuthSheetsScopeException : Exception
    {
        /// <summary>
        /// Initializes a new OAuthSheetsScopeException instance.
        /// </summary>
        public OAuthSheetsScopeException() { }

        /// <summary>
        /// Initializes a new OAuthSheetsScopeException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public OAuthSheetsScopeException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new OAuthSheetsScopeException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public OAuthSheetsScopeException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected OAuthSheetsScopeException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}
