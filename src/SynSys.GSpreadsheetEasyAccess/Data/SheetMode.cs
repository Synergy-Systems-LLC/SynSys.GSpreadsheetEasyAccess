namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Перечисление для определённого заполнения листа.
    /// Выбор режима влияет на состояние ячеек и строк листа.
    /// </summary>
    public enum SheetMode
    {
        /// <summary>
        /// Таблица без шапки и ключа
        /// </summary>
        Simple,
        /// <summary>
        /// Таблица с шапкой в одну строку
        /// </summary>
        Head,
        /// <summary>
        /// Таблица с шапкой в одну строку и столбцом-ключём
        /// </summary>
        HeadAndKey
    }
}
