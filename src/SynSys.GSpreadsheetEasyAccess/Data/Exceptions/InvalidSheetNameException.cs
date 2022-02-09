using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за некорректного имени листа гугл таблицы
    /// </summary>
    [Serializable]
    public class InvalidSheetNameException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetNameException
        /// </summary>
        public InvalidSheetNameException() { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetNameException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public InvalidSheetNameException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetNameException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public InvalidSheetNameException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected InvalidSheetNameException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}