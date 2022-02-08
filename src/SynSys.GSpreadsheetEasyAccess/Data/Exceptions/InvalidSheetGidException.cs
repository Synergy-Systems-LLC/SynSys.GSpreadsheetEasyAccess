using System;

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
        public InvalidSheetGidException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetGidException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        public InvalidSheetGidException(string message, Exception innerException) : base(message, innerException) { }
    }
}