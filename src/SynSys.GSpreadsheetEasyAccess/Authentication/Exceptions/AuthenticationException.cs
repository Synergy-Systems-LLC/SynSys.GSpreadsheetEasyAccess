using System;

namespace SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее во время аутентификации
    /// </summary>
    [Serializable]
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
