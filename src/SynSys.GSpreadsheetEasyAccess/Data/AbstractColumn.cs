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
        /// Column number, not index.
        /// </summary>
        public int Number { get; protected set; }

        /// <summary>
        /// Title of column.
        /// </summary>
        public string Title { get; protected set; }

        /// <summary>
        /// Current status.
        /// </summary>
        public ChangeStatus Status { get; internal set; } = ChangeStatus.Original;

        /// <summary>
        /// Change column number.
        /// </summary>
        /// <param name="newNumber">Not index</param>
        public abstract void ChangeNumber(int newNumber);

        /// <summary>
        /// Determines whether the specified object is equal to the current object.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public abstract bool Equals(AbstractColumn other);
    }
}