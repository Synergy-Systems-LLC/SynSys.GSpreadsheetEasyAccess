using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за несоответствия областей действия приложения и отправляемых запросов
    /// </summary>
    [Serializable]
    public class OAuthSheetsScopeException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс OAuthSheetsScopeException
        /// </summary>
        public OAuthSheetsScopeException() { }

        /// <summary>
        /// Инициализирует новый инстанс OAuthSheetsScopeException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public OAuthSheetsScopeException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс OAuthSheetsScopeException с сообщением об ошибке и ссылкой на причину текущего исключения
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
