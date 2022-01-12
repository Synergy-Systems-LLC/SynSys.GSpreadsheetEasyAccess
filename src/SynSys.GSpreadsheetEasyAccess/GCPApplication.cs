using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SynSys.GSpreadsheetEasyAccess
{
    class GCPApplication
    {
        private SheetsService _sheetsService;
        private Principal _principal;
 
        /// <summary>
        /// Имя приложения, которое будет использоваться в заголовке User-Agent.<br/>
        /// Значение по умолчанию <c>string.Empty</c>. 
        /// </summary>
        public string ApplicationName { get; set; } = string.Empty;

        /// <summary>
        /// Для получения доступа к сервису гугл таблиц необходимо пройти аутентификацию.
        /// Необходимо указать кто аутентифицируется.
        /// </summary>
        /// <param name="principal"></param>
        public void AuthenticateAs(Principal principal)
        {
            _principal = principal;
            _sheetsService = principal.GetSheetsService();
        }

        public SheetModel GetSheet(string uri)
        {
            CheckUri(uri);

            return GetSheet(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri)
            );
        }

        public SheetModel GetSheet(string spreadsheetId, int gid)
        {
            CheckSheetAttribut(gid);

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            ValidateData(data);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.Simple, string.Empty, data);
        }

        public SheetModel GetSheet(string spreadsheetId, string sheetName)
        {
            CheckSheetAttribut(sheetName);

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);
            IList<IList<object>> data = GetData(spreadsheetId, sheetName);

            ValidateData(data);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.Simple, string.Empty, data);
        }

        public SheetModel GetSheetWithHead(string uri)
        {
            CheckUri(uri);

            return GetSheetWithHead(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri)
            );
        }

        public SheetModel GetSheetWithHead(string spreadsheetId, int gid)
        {
            CheckSheetAttribut(gid);

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            ValidateData(data);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.Head, string.Empty, data);
        }

        public SheetModel GetSheetWithHead(string spreadsheetId, string sheetName)
        {
            CheckSheetAttribut(sheetName);

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);
            IList<IList<object>> data = GetData(spreadsheetId, sheetName);

            ValidateData(data);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.Head, string.Empty, data);
        }

        public SheetModel GetSheetWithHeadAndKey(string uri, string keyName)
        {
            CheckUri(uri);

            return GetSheetWithHeadAndKey(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri),
                keyName
            );
        }

        public SheetModel GetSheetWithHeadAndKey(string spreadsheetId, int gid, string keyName)
        {
            CheckSheetAttribut(gid);

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            ValidateData(data, keyName);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.HeadAndKey, keyName, data);
        }

        public SheetModel GetSheetWithHeadAndKey(string spreadsheetId, string sheetName, string keyName)
        {
            CheckSheetAttribut(sheetName);

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);
            IList<IList<object>> data = GetData(spreadsheetId, sheetName);

            ValidateData(data, keyName);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.HeadAndKey, keyName, data);
        }

        public void UpdateSheet(SheetModel sheetModel)
        {
            if (_principal == null)
            {
                throw new InvalidOperationException(
                    $"Перед тем как обновлять лист, необходимо аутентифицироваться."
                );
            }

            if (_principal is ServiceAccount)
            {
                throw new InvalidOperationException(
                    $"Используя {nameof(ServiceAccount)} нельзя обновлять листы гугл таблиц."
                );
            }

            CreateAppendRequest(sheetModel)?.Execute();
            CreateUpdateRequest(sheetModel)?.Execute();
            CreateDeleteRequest(sheetModel)?.Execute();

            sheetModel.ClearDeletedRows();
            sheetModel.RenumberRows();
            sheetModel.ResetRowStatuses();
        }


        #region GetSheetModel
        private void CheckUri(string uri)
        {
            if (string.IsNullOrWhiteSpace(uri))
            {
                throw new ArgumentException("uri не может быть пустым");
            }
        }

        private void CheckSheetAttribut(int gid)
        {
            if (gid < 0)
            {
                throw new ArgumentException(
                    $"Был введён некорректный gid листа. Возможно он отсутствует в uri"
                );
            }
        }

        private void CheckSheetAttribut(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException(
                    $"Не существует листа с пустым именем" // Подумай над формулировкой
                );
            }
        }

        private Spreadsheet GetGoogleSpreadsheet(string spreadsheetId)
        {
            return _sheetsService.Spreadsheets.Get(spreadsheetId).Execute();
        }

        private IList<IList<object>> GetData(string spreadsheetId, string sheetName)
        {
            return _sheetsService
                .Spreadsheets
                .Values
                .Get(spreadsheetId, sheetName)
                .Execute()
                .Values;
        }

        private void ValidateData(IList<IList<object>> data)
        {
            if (data == null)
            {
                throw new InvalidOperationException($"Лист не содержит данных");
            }
        }

        private void ValidateData(IList<IList<object>> data, string keyName)
        {
            ValidateData(data);

            if (!data[0].Contains(keyName))
            {
                throw new InvalidOperationException($"Лист не содержит ключ {keyName}");
            }
        }

        private SheetModel CreateSheetModel(Spreadsheet spreadsheet, Sheet sheet, SheetMode mode, string keyName, IList<IList<object>> data)
        {
            var sheetModel = new SheetModel()
            {
                SpreadsheetId = spreadsheet.SpreadsheetId,
                SpreadsheetTitle = spreadsheet.Properties.Title,
                Gid = sheet.Properties.SheetId.ToString(),
                Title = sheet.Properties.Title,
                Mode = mode,
                KeyName = keyName
            };

            sheetModel.Fill(data);

            return sheetModel;
        }
        #endregion

        #region UpdateSheetModel
        /// <summary>
        /// Создание запроса на обновление листа в google таблице.
        /// В запросе содержатся изменения значений в ячейках листа.
        /// </summary>
        /// <param name="sheet">Экземпляр изменённого листа</param>
        /// <returns></returns>
        private SpreadsheetsResource.ValuesResource.BatchUpdateRequest CreateUpdateRequest(SheetModel sheet)
        {
            if (sheet.Rows.FindAll(row => row.Status == RowStatus.ToChange).Count <= 0) return null;
 
            var requestBody = new BatchUpdateValuesRequest
            {
                Data = sheet.GetChangeValueRange(),
                ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum
                    .RAW
                    .ToString()
            };

            return _sheetsService
                .Spreadsheets
                .Values
                .BatchUpdate(requestBody, sheet.SpreadsheetId);
        }

        /// <summary>
        /// Создание запроса на обновление листа в google таблице.
        /// В запросе содержатся группы строк для удаления.
        /// </summary>
        /// <param name="sheet">Экземпляр изменённого листа</param>
        /// <returns></returns>
        private SpreadsheetsResource.BatchUpdateRequest CreateDeleteRequest(SheetModel sheet)
        {
            if (sheet.Rows.FindAll(row => row.Status == RowStatus.ToDelete).Count > 0)
            {
                var requestBody = new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request>()
                };

                foreach (List<Row> groupRows in sheet.GetDeleteRows())
                {
                    requestBody.Requests.Add(
                        CreateDeleteDimensionRequest(
                            Convert.ToInt32(sheet.Gid),
                            groupRows.Last().Number - 1,
                            groupRows.First().Number
                        )
                    );
                }

                return new SpreadsheetsResource.BatchUpdateRequest(
                    _sheetsService,
                    requestBody,
                    sheet.SpreadsheetId
                );
            }

            return null;
        }

        /// <summary>
        /// Создание запроса на удаление группы строк.
        /// </summary>
        /// <param name="gid">Id листа</param>
        /// <param name="startRow">Начальная строка из группы строк для удаления</param>
        /// <param name="endRow">Конечная строка из группы строк для удаления</param>
        /// <returns></returns>
        private Request CreateDeleteDimensionRequest(int gid, int startRow, int endRow)
        {
            return new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        SheetId = gid,
                        Dimension = "ROWS",
                        StartIndex = startRow,
                        EndIndex = endRow,
                    }
                }
            };
        }

        /// <summary>
        /// Создание запроса на обновление листа в google таблице.
        /// В запросе содержатся строки для добавления в конец листа.
        /// </summary>
        /// <param name="sheet">Экземпляр изменённого листа</param>
        /// <returns></returns>
        private SpreadsheetsResource.ValuesResource.AppendRequest CreateAppendRequest(SheetModel sheet)
        {
            if (sheet.Rows.FindAll(row => row.Status == RowStatus.ToAppend).Count <= 0) return null;
 
            var request = _sheetsService
                .Spreadsheets
                .Values
                .Append(sheet.GetAppendValueRange(), sheet.SpreadsheetId, $"{sheet.Title}!A:A");

            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum
                .USERENTERED;

            return request;
        }
        #endregion
    }
}
