using System;

namespace SynSys.GSpreadsheetEasyAccess.Data
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
        public InvalidSheetNameException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс InvalidSheetNameException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        public InvalidSheetNameException(string message, Exception innerException) : base(message, innerException) { }
    }
}