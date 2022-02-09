using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за некорректного id листа гугл таблицы
    /// </summary>
    [Serializable]
    public class InvalidSheetGidException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetGidException
        /// </summary>
        public InvalidSheetGidException() { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetGidException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public InvalidSheetGidException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetGidException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public InvalidSheetGidException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected InvalidSheetGidException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}