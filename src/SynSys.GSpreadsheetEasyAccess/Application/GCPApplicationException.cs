using System;

namespace SynSys.GSpreadsheetEasyAccess.Application
{
    /// <summary>
    /// Представляет исключение возникшее в классе GCPApplication
    /// </summary>
    [Serializable]
    public class GCPApplicationException : Exception
    {
        /// <summary>
        /// Инициализирует новый инстанс GCPApplicationException
        /// </summary>
        public GCPApplicationException() : base() { }

        /// <summary>
        /// Инициализирует новый инстанс GCPApplicationException с сообщением об ошибке
        /// </summary>
        public GCPApplicationException(string message) : base(message) { }

        /// <summary>
        /// Инициализирует новый инстанс GCPApplicationException с сообщением об ошибке и ссылкой на причину текущего исключения
        /// </summary>
        public GCPApplicationException(string message, Exception innerException) : base(message, innerException) { }
    }
}
