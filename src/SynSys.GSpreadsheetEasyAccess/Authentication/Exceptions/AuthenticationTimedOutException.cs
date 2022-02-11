using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за вышедшего времени на аутентификацию
    /// </summary>
    [Serializable]
    public class AuthenticationTimedOutException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс AuthenticationTimedOutException
        /// </summary>
        public AuthenticationTimedOutException() { }

        /// <summary>
        /// Инициализирует новый инстанс AuthenticationTimedOutException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public AuthenticationTimedOutException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс AuthenticationTimedOutException с сообщением об ошибке и ссылкой на причину текущего исключения
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