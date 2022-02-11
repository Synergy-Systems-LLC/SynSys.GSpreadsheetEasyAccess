using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за отсутствия хотя бы одного требуемого столбца
    /// </summary>
    [Serializable]
    public class InvalidSheetHeadException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetHeadException
        /// </summary>
        public InvalidSheetHeadException() { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetHeadException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public InvalidSheetHeadException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetHeadException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public InvalidSheetHeadException(string message, Exception innerException) : base(message, innerException) { }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected InvalidSheetHeadException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// Лист, состояние которого вызвало исключение
        /// </summary>
        public SheetModel Sheet { get; internal set; }

        /// <summary>
        /// Требуемые заголовки столбцов, которые не были найдены в листе
        /// </summary>
        public IEnumerable<string> LostedHeaders { get; internal set; }
    }
}