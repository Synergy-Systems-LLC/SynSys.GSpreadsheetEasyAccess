using System;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    public class AuthenticationException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс AuthenticationException
        /// </summary>
        public AuthenticationException() : base() { }

        /// <summary>
        /// Инициализирует новый инстанс AuthenticationException с сообщением об ошибке
        /// </summary>
        public AuthenticationException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс AuthenticationException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        public AuthenticationException(string message, Exception innerException) : base(message, innerException) { }
    }
}
