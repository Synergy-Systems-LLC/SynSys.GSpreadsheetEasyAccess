using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за отсутствия ключевого столбца
    /// </summary>
    [Serializable]
    public class SheetKeyNotFoundException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс SheetKeyNotFoundException
        /// </summary>
        public SheetKeyNotFoundException() { }

        /// <summary>
        /// Инициализирует новый инстанс SheetKeyNotFoundException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public SheetKeyNotFoundException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс SheetKeyNotFoundException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SheetKeyNotFoundException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected SheetKeyNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        /// <summary>
        /// Лист, состояние которого вызвало исключение
        /// </summary>
        public SheetModel Sheet { get; set; }
    }
}