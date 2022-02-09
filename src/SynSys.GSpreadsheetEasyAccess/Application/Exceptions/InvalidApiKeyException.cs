using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за невалидного api key приложения на Google Cloud Platform
    /// </summary>
    [Serializable]
    public class InvalidApiKeyException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс InvalidApiKeyException
        /// </summary>
        public InvalidApiKeyException() { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidApiKeyException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public InvalidApiKeyException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidApiKeyException с сообщением об ошибке и ссылкой на причину текущего исключения
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