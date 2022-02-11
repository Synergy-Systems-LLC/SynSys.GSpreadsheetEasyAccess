using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за того что данному пользователю закрыт доступ 
    /// к каким либо действиям с google таблицами
    /// </summary>
    [Serializable]
    public class UserAccessDeniedException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс UserAccessDeniedException
        /// </summary>
        public UserAccessDeniedException() { }

        /// <summary>
        /// Инициализирует новый инстанс UserAccessDeniedException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public UserAccessDeniedException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс UserAccessDeniedException с сообщением об ошибке и ссылкой на причину текущего исключения
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
        /// Содержит операцию к которой нет доступа у пользователя
        /// </summary>
        public string Operation { get; internal set; }
    }
}