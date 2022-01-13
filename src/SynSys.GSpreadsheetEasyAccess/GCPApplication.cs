using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SynSys.GSpreadsheetEasyAccess
{
    /// <summary>
    /// Задача класса получать и обновлять данные из гугл таблиц.
    /// </summary>
    /// <remarks>
    /// Пользоваться методами класса можно только после успешной аутентификации.
    /// </remarks>
    class GCPApplication
    {
        private SheetsService _sheetsService;
        private Principal _principal;
 
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

        /// <summary>
        /// Получение данных из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представлен списком строк.<br/>
        /// Шапка отсутствует.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение.
        /// </summary>
        /// <param name="uri">Полный uri адрес листа</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public SheetModel GetSheet(string uri)
        {
            CheckPrincipal();
            CheckUri(uri);

            return GetSheet(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri)
            );
        }

        /// <summary>
        /// Получение данных из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представлен списком строк.<br/>
        /// Шапка отсутствует.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение.
        /// </summary>
        /// <param name="spreadsheetId">Id таблицы</param>
        /// <param name="gid">Id листа</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public SheetModel GetSheet(string spreadsheetId, int gid)
        {
            CheckPrincipal();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            ValidateData(data);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.Simple, string.Empty, data);
        }

        /// <summary>
        /// Получение данных из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представлен списком строк.<br/>
        /// Шапка отсутствует.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение.
        /// </summary>
        /// <param name="spreadsheetId">Id таблицы</param>
        /// <param name="sheetName">Имя листа</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public SheetModel GetSheet(string spreadsheetId, string sheetName)
        {
            CheckPrincipal();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);
            IList<IList<object>> data = GetData(spreadsheetId, sheetName);

            ValidateData(data);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.Simple, string.Empty, data);
        }

        /// <summary>
        /// Получение данных из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представлен списком строк, без первой строки.<br/>
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="uri">Полный uri адрес листа</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public SheetModel GetSheetWithHead(string uri)
        {
            CheckPrincipal();
            CheckUri(uri);

            return GetSheetWithHead(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri)
            );
        }

        /// <summary>
        /// Получение данных из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представлен списком строк, без первой строки.<br/>
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="spreadsheetId">Id таблицы</param>
        /// <param name="gid">Id листа</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public SheetModel GetSheetWithHead(string spreadsheetId, int gid)
        {
            CheckPrincipal();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            ValidateData(data);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.Head, string.Empty, data);
        }

        /// <summary>
        /// Получение данных из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представлен списком строк, без первой строки.<br/>
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="spreadsheetId">Id таблицы</param>
        /// <param name="sheetName">Имя листа</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public SheetModel GetSheetWithHead(string spreadsheetId, string sheetName)
        {
            CheckPrincipal();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);
            IList<IList<object>> data = GetData(spreadsheetId, sheetName);

            ValidateData(data);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.Head, string.Empty, data);
        }

        /// <summary>
        /// Получение данных из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представлен списком строк, без первой строки.<br/>
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек и есть ключевая ячейка.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="uri"></param>
        /// <param name="keyName"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public SheetModel GetSheetWithHeadAndKey(string uri, string keyName)
        {
            CheckPrincipal();
            CheckUri(uri);

            return GetSheetWithHeadAndKey(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri),
                keyName
            );
        }

        /// <summary>
        /// Получение данных из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представлен списком строк, без первой строки.<br/>
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек и есть ключевая ячейка.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="gid"></param>
        /// <param name="keyName"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public SheetModel GetSheetWithHeadAndKey(string spreadsheetId, int gid, string keyName)
        {
            CheckPrincipal();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            ValidateData(data, keyName);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.HeadAndKey, keyName, data);
        }

        /// <summary>
        /// Получение данных из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представлен списком строк, без первой строки.<br/>
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек и есть ключевая ячейка.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetName"></param>
        /// <param name="keyName"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public SheetModel GetSheetWithHeadAndKey(string spreadsheetId, string sheetName, string keyName)
        {
            CheckPrincipal();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);
            IList<IList<object>> data = GetData(spreadsheetId, sheetName);

            ValidateData(data, keyName);

            return CreateSheetModel(spreadsheet, sheet, SheetMode.HeadAndKey, keyName, data);
        }

        /// <summary>
        /// Обновленние листа гугл таблицы на основе изменённого экземпляра типа SheetModel.
        /// </summary>
        /// <remarks>
        /// Метод изменяет данные в ячейках,
        /// добавляет строки в конец листа и удаляет выбраные строки.<br />
        /// Все эти действия просходят на основе запросов в google.
        /// </remarks>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        public void UpdateSheet(SheetModel sheetModel)
        {
            CheckPrincipal();

            CreateAppendRequest(sheetModel)?.Execute();
            CreateUpdateRequest(sheetModel)?.Execute();
            CreateDeleteRequest(sheetModel)?.Execute();

            sheetModel.ClearDeletedRows();
            sheetModel.RenumberRows();
            sheetModel.ResetRowStatuses();
        }


        #region GetSheetModel
        private void CheckPrincipal()
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
        }

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
