using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SynSys.GSpreadsheetEasyAccess
{
    /// <summary>
    /// Класс представляет приложение на Google Cloud Platform у которого есть доступ
    /// к сервису google таблиц 
    /// <a href="https://developers.google.com/sheets/api?hl=en_US">Google Sheets API</a>.
    /// </summary>
    /// <remarks>
    /// Задача класса получать и обновлять данные из google таблиц.<br/>
    /// Пользоваться методами можно только после успешной аутентификации.
    /// </remarks>
    public class GCPApplication
    {
        private SheetsService _sheetsService;
        private Principal _principal;

        /// <summary>
        /// Для получения доступа к сервису гугл таблиц необходимо пройти аутентификацию.
        /// Необходимо указать кто аутентифицируется.
        /// </summary>
        /// <param name="principal"></param>
        /// <exception cref="GCPApplicationException"></exception>
        public void AuthenticateAs(Principal principal)
        {
            try
            {
                _principal = principal;
                _sheetsService = principal.GetSheetsService();
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не получилось авторизовать пользователя", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public SheetModel GetSheet(string uri)
        {
            try
            {
                CheckPrincipal();
                SheetModel.CheckUri(uri);

                return GetSheet(
                    HttpUtils.GetSpreadsheetIdFromUri(uri),
                    HttpUtils.GetGidFromUri(uri)
                );
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось создать SheetModel", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public SheetModel GetSheet(string spreadsheetId, int gid)
        {
            try
            {
                CheckPrincipal();
                SheetModel.CheckGid(gid);

                Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
                Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);
                IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

                SheetModel.ValidateData(data);

                return SheetModel.Create(spreadsheet, sheet, SheetMode.Simple, string.Empty, data);
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось создать SheetModel", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public SheetModel GetSheet(string spreadsheetId, string sheetName)
        {
            try 
            {
                CheckPrincipal();
                SheetModel.CheckName(sheetName);

                Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
                Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);
                IList<IList<object>> data = GetData(spreadsheetId, sheetName);

                SheetModel.ValidateData(data);

                return SheetModel.Create(spreadsheet, sheet, SheetMode.Simple, string.Empty, data);
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось создать SheetModel", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public SheetModel GetSheetWithHead(string uri)
        {
            try
            {
                CheckPrincipal();
                SheetModel.CheckUri(uri);

                return GetSheetWithHead(
                    HttpUtils.GetSpreadsheetIdFromUri(uri),
                    HttpUtils.GetGidFromUri(uri)
                );
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось создать SheetModel", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public SheetModel GetSheetWithHead(string spreadsheetId, int gid)
        {
            try
            {
                CheckPrincipal();
                SheetModel.CheckGid(gid);

                Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
                Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);
                IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

                SheetModel.ValidateData(data);

                return SheetModel.Create(spreadsheet, sheet, SheetMode.Head, string.Empty, data);
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось создать SheetModel", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public SheetModel GetSheetWithHead(string spreadsheetId, string sheetName)
        {
            try
            {
                CheckPrincipal();
                SheetModel.CheckName(sheetName);

                Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
                Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);
                IList<IList<object>> data = GetData(spreadsheetId, sheetName);

                SheetModel.ValidateData(data);

                return SheetModel.Create(spreadsheet, sheet, SheetMode.Head, string.Empty, data);
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось создать SheetModel", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public SheetModel GetSheetWithHeadAndKey(string uri, string keyName)
        {
            try
            {
                CheckPrincipal();
                SheetModel.CheckUri(uri);

                return GetSheetWithHeadAndKey(
                    HttpUtils.GetSpreadsheetIdFromUri(uri),
                    HttpUtils.GetGidFromUri(uri),
                    keyName
                );
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось создать SheetModel", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public SheetModel GetSheetWithHeadAndKey(string spreadsheetId, int gid, string keyName)
        {
            try
            {
                CheckPrincipal();
                SheetModel.CheckGid(gid);

                Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
                Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);
                IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

                SheetModel.ValidateData(data, keyName);

                return SheetModel.Create(spreadsheet, sheet, SheetMode.HeadAndKey, keyName, data);
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось создать SheetModel", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public SheetModel GetSheetWithHeadAndKey(string spreadsheetId, string sheetName, string keyName)
        {
            try
            {
                CheckPrincipal();
                SheetModel.CheckName(sheetName);

                Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
                Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);
                IList<IList<object>> data = GetData(spreadsheetId, sheetName);

                SheetModel.ValidateData(data, keyName);

                return SheetModel.Create(spreadsheet, sheet, SheetMode.HeadAndKey, keyName, data);
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось создать SheetModel", e);
            }
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
        /// <exception cref="GCPApplicationException"></exception>
        public void UpdateSheet(SheetModel sheetModel)
        {
            try
            {
                CheckPrincipalForUpdateSheet();

                CreateAppendRequest(sheetModel)?.Execute();
                CreateUpdateRequest(sheetModel)?.Execute();
                CreateDeleteRequest(sheetModel)?.Execute();

                sheetModel.ClearDeletedRows();
                sheetModel.RenumberRows();
                sheetModel.ResetRowStatuses();
            }
            catch(Exception e)
            {
                throw new GCPApplicationException("Не удалось обновить данные на google spreadsheet", e);
            }
        }


        #region CheckPrincipal
        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        private void CheckPrincipal()
        {
            if (_principal == null)
            {
                throw new InvalidOperationException(
                    $"Перед тем как обновлять лист, необходимо аутентифицироваться."
                );
            }
        }

        /// <exception cref="InvalidOperationException">
        /// Если principal null или ServiceAccount
        /// </exception>
        private void CheckPrincipalForUpdateSheet()
        {
            CheckPrincipal();

            if (_principal is ServiceAccount)
            {
                throw new InvalidOperationException(
                    $"Используя {nameof(ServiceAccount)} нельзя обновлять листы гугл таблиц."
                );
            }
        }
        #endregion

        #region Spreadsheets
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
