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

        /// <summary>
        /// Имя таблицы в которой искали лист
        /// </summary>
        public string SpreadsheetName { get; internal set; }

        /// <summary>
        /// Id таблицы в которой искали лист
        /// </summary>
        public string SpreadsheetId { get; internal set; }

        /// <summary>
        /// Id листа по которому не удалось найти лист
        /// </summary>
        public string SheetGid { get; internal set; }

        /// <summary>
        /// Имя листа по которому не удалось найти лист
        /// </summary>
        public string SheetName { get; internal set; }
    }
}