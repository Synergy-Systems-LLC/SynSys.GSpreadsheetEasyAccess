using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за некорректного Uri листа google таблицы
    /// </summary>
    [Serializable]
    public class InvalidSheetUriException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetUriException
        /// </summary>
        public InvalidSheetUriException() { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetUriException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public InvalidSheetUriException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetUriException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public InvalidSheetUriException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected InvalidSheetUriException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}