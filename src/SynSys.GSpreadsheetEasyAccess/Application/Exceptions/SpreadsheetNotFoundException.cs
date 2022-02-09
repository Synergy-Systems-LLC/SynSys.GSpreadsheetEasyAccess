using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за того что не получилось найти таблицу по текущему spreadsheet id
    /// </summary>
    [Serializable]
    public class SpreadsheetNotFoundException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс SpreadsheetNotFoundException
        /// </summary>
        public SpreadsheetNotFoundException() { }

        /// <summary>
        /// Инициализирует новый инстанс SpreadsheetNotFoundException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public SpreadsheetNotFoundException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс SpreadsheetNotFoundException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SpreadsheetNotFoundException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected SpreadsheetNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}