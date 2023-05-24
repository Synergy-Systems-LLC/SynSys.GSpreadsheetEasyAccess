using System;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Represents an abstract column of any table.
    /// </summary>
    /// <remarks>
    /// Does not have cells or data from the table.
    /// </remarks>
    public abstract class AbstractColumn : IEquatable<AbstractColumn>
    {
        /// <summary>
        /// 
        /// </summary>
        protected int _number;

        /// <summary>
        /// 
        /// </summary>
        protected string _title;

        /// <summary>
        /// Column number, not index.
        /// </summary>
        public int Number => _number;


        /// <summary>
        /// Title of column.
        /// </summary>
        public string Title => _title;

        /// <summary>
        /// Current status.
        /// </summary>
        public RowStatus Status { get; internal set; } = RowStatus.Original;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(AbstractColumn other)
        {
            // В сравнении нет Number потому что главным определяющим параметром является заголовок.
            // Столбец может находится на любой позиции, но имя из-за этого меняться не будет.
            return _title == other.Title;
        }
    }
}