using Newtonsoft.Json;

namespace SynSys.GSpreadsheetEasyAccess
{
    /// <summary>
    /// Предоставляет методы для конвертации экземпляра SheetModel 
    /// в строковое представление JSON и обратно.
    /// </summary>
    public class JsonSerialization
    {
        /// <summary>
        /// Сериализует экземпляр указаного листа в JSON с применением выбранного форматирования.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="formatting"></param>
        /// <returns>Строковое представление объекта в формате JSON</returns>
        public static string SerializeSheet(SheetModel sheet, Formatting formatting)
        {
            // PreserveReferencesHandling.Objects - обязательное свойство так как в Cell есть
            // ссылка на Row.
            return JsonConvert.SerializeObject(
                sheet,
                formatting,
                new JsonSerializerSettings()
                {
                    PreserveReferencesHandling = PreserveReferencesHandling.Objects
                }
            );
        }

        /// <summary>
        /// Десериализует JSON до экземпляра SheetModel.
        /// </summary>
        /// <param name="jsonSheet"></param>
        /// <returns>Десериализованный объект из строки JSON</returns>
        public static SheetModel DeserializeSheet(string jsonSheet)
        {
            return JsonConvert.DeserializeObject<SheetModel>(jsonSheet);
        }
    }
}
