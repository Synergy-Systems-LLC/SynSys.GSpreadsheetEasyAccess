using Newtonsoft.Json;

namespace SynSys.GSpreadsheetEasyAccess.Data
{
    /// <summary>
    /// Provides methods for converting a SheetModel instance 
    /// to and from a JSON string representation.
    /// </summary>
    public class JsonSerialization
    {
        /// <summary>
        /// Serializes an instance of the specified sheet to JSON using the selected formatting.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="formatting"></param>
        /// <returns>The string representation of the object in the format JSON.</returns>
        public static string SerializeSheet(SheetModel sheet, Formatting formatting)
        {
            // PreserveReferencesHandling.Objects - required property since Cell has link to Row.
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
        /// Deserializes JSON to an instance SheetModel.
        /// </summary>
        /// <param name="jsonSheet"></param>
        /// <returns>Deserialized object from string JSON.</returns>
        public static SheetModel DeserializeSheet(string jsonSheet)
        {
            return JsonConvert.DeserializeObject<SheetModel>(jsonSheet);
        }
    }
}
