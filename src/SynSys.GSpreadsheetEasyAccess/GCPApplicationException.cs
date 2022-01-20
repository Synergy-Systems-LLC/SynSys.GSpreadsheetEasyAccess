using System;

namespace SynSys.GSpreadsheetEasyAccess
{
    /// <summary>
    /// Представляет исключение возникшее в классе GCPApplication
    /// </summary>
    public class GCPApplicationException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс GCPApplication
        /// </summary>
        public GCPApplicationException() : base() { }

        /// <summary>
        /// Инициализирует новый инстанс GCPApplication с сообщением об ошибке
        /// </summary>
        public GCPApplicationException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс GCPApplication с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        public GCPApplicationException(string message, Exception innerException) : base(message, innerException) { }
    }
}
