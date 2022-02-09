using System;
using System.Runtime.Serialization;

namespace SynSys.GSpreadsheetEasyAccess.Application.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за того что не получилось найти лист по текущему sheet id
    /// </summary>
    [Serializable]
    public class SheetNotFoundException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс SheetNotFoundException
        /// </summary>
        public SheetNotFoundException() { }

        /// <summary>
        /// Инициализирует новый инстанс SheetNotFoundException с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public SheetNotFoundException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс SheetNotFoundException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public SheetNotFoundException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected SheetNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}