using System;

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
        public EmptySheetException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс EmptySheetException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        public EmptySheetException(string message, Exception innerException) : base(message, innerException) { }
    }
}