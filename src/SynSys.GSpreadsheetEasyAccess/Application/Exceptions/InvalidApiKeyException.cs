using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Represents an exception throwndue to an invalid API key of the application on the Google Cloud Platform.
    /// </summary>
    [Serializable]
    public class InvalidApiKeyException : Exception
    {
        /// <summary>
        /// Initializes a new InvalidApiKeyException instance.
        /// </summary>
        public InvalidApiKeyException() { }

        /// <summary>
        /// Initializes a new InvalidApiKeyException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public InvalidApiKeyException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new InvalidApiKeyException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public InvalidApiKeyException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected InvalidApiKeyException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}