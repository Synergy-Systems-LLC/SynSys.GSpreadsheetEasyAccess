using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за того что пользовать отменил процесс аутентификации
    /// </summary>
    [Serializable]
    public class UserCanceledAuthenticationException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс UserCanceledAuthenticationException
        /// </summary>
        public UserCanceledAuthenticationException() { }

        /// <summary>
        /// Инициализирует новый инстанс UserCanceledAuthenticationException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public UserCanceledAuthenticationException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс UserCanceledAuthenticationException с сообщением об ошибке и ссылкой на причину текущего исключения
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