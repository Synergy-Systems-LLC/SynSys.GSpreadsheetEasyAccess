using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Represents an exception thrown due to the given user has been denied access
    /// to any actions with Google spreadsheet.
    /// </summary>
    [Serializable]
    public class UserAccessDeniedException : Exception
    {
        /// <summary>
        /// Initializes a new UserAccessDeniedException instance.
        /// </summary>
        public UserAccessDeniedException() { }

        /// <summary>
        /// Initializes a new UserAccessDeniedException instance with a message about exception.
        /// </summary>
        /// <param name="message"></param>
        public UserAccessDeniedException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new UserAccessDeniedException instance with an error message and a reference to the reason for the current exception.
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public UserAccessDeniedException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected UserAccessDeniedException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// Contains the name of the operation to which the user does not have access.
        /// </summary>
        public string Operation { get; internal set; }
    }
}