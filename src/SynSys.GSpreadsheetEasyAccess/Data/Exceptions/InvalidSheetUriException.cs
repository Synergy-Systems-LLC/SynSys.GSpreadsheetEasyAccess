using System;

namespace SynSys.GSpreadsheetEasyAccess.Data.Exceptions
{
    /// <summary>
    /// Представляет исключение возникшее из-за некорректного Uri листа гугл таблицы
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
        public InvalidSheetUriException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetUriException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        public InvalidSheetUriException(string message, Exception innerException) : base(message, innerException) { }
    }
}