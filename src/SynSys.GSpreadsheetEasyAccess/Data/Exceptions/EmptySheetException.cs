using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за некорректных данных в листе гугл таблицы
    /// </summary>
    [Serializable]
    public class EmptySheetException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс EmptySheetException
        /// </summary>
        public EmptySheetException() { }

        /// <summary>
        /// Инициализирует новый инстанс EmptySheetException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public EmptySheetException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс EmptySheetException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public EmptySheetException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected EmptySheetException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// Лист, состояние которого вызвало исключение
        /// </summary>
        public SheetModel Sheet { get; set; }
    }
}