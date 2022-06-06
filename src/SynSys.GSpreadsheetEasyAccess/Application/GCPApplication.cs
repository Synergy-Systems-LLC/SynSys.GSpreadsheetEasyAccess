using Google;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using SynSys.GSpreadsheetEasyAccess.Application.Exceptions;
using SynSys.GSpreadsheetEasyAccess.Authentication;
using SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions;
using SynSys.GSpreadsheetEasyAccess.Data;
using SynSys.GSpreadsheetEasyAccess.Data.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace SynSys.GSpreadsheetEasyAccess.Application
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
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="AuthenticationTimedOutException"></exception>
        /// <exception cref="UserCanceledAuthenticationException"></exception>
        public void AuthenticateAs(Principal principal)
        {
            _principal = principal ?? throw new ArgumentNullException(nameof(principal));
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        public SheetModel GetSheet(string uri)
        {
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        public SheetModel GetSheet(string spreadsheetId, int gid)
        {
            CheckSheetService();

            var sheetModel = new SheetModel()
            {
                SpreadsheetId = spreadsheetId,
                Gid = gid
            };

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            sheetModel.SpreadsheetTitle = spreadsheet.Properties.Title;
            sheetModel.Title = sheet.Properties.Title;
            sheetModel.Mode = SheetMode.Simple;
            sheetModel.KeyName = string.Empty;
            sheetModel.Fill(data);

            return sheetModel;
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        public SheetModel GetSheet(string spreadsheetId, string sheetName)
        {
            CheckSheetService();

            var sheetModel = new SheetModel()
            {
                SpreadsheetId = spreadsheetId,
                Title = sheetName
            };

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, sheetName);
            IList<IList<object>> data = GetData(spreadsheetId, sheetName);

            sheetModel.SpreadsheetTitle = spreadsheet.Properties.Title;
            sheetModel.Gid = sheet.Properties.SheetId.Value;
            sheetModel.Mode = SheetMode.Simple;
            sheetModel.KeyName = string.Empty;
            sheetModel.Fill(data);

            return sheetModel;
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        /// <exception cref="EmptySheetException"></exception>
        public SheetModel GetSheetWithHead(string uri)
        {
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        /// <exception cref="EmptySheetException"></exception>
        public SheetModel GetSheetWithHead(string spreadsheetId, int gid)
        {
            CheckSheetService();

            var sheetModel = new SheetModel()
            {
                SpreadsheetId = spreadsheetId,
                Gid = gid
            };

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            sheetModel.SpreadsheetTitle = spreadsheet.Properties.Title;
            sheetModel.Title = sheet.Properties.Title;
            sheetModel.Mode = SheetMode.Head;
            sheetModel.KeyName = string.Empty;

            sheetModel.ValidateData(data);
            sheetModel.Fill(data);

            return sheetModel;
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        /// <exception cref="EmptySheetException"></exception>
        public SheetModel GetSheetWithHead(string spreadsheetId, string sheetName)
        {
            CheckSheetService();

            var sheetModel = new SheetModel()
            {
                SpreadsheetId = spreadsheetId,
                Title = sheetName
            };

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, sheetName);
            IList<IList<object>> data = GetData(spreadsheetId, sheetName);

            sheetModel.SpreadsheetTitle = spreadsheet.Properties.Title;
            sheetModel.Gid = sheet.Properties.SheetId.Value;
            sheetModel.Mode = SheetMode.Head;
            sheetModel.KeyName = string.Empty;

            sheetModel.ValidateData(data);
            sheetModel.Fill(data);

            return sheetModel;
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        /// <exception cref="SheetKeyNotFoundException"></exception>
        /// <exception cref="EmptySheetException"></exception>
        /// <returns></returns>
        public SheetModel GetSheetWithHeadAndKey(string uri, string keyName)
        {
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        /// <exception cref="SheetKeyNotFoundException"></exception>
        /// <exception cref="EmptySheetException"></exception>
        public SheetModel GetSheetWithHeadAndKey(string spreadsheetId, int gid, string keyName)
        {
            CheckSheetService();

            var sheetModel = new SheetModel()
            {
                SpreadsheetId = spreadsheetId,
                Gid = gid
            };

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            sheetModel.SpreadsheetTitle = spreadsheet.Properties.Title;
            sheetModel.Title = sheet.Properties.Title;
            sheetModel.Mode = SheetMode.HeadAndKey;
            sheetModel.KeyName = keyName;

            sheetModel.ValidateData(data, keyName);
            sheetModel.Fill(data);

            return sheetModel;
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        /// <exception cref="SheetKeyNotFoundException"></exception>
        /// <exception cref="EmptySheetException"></exception>
        public SheetModel GetSheetWithHeadAndKey(string spreadsheetId, string sheetName, string keyName)
        {
            CheckSheetService();

            var sheetModel = new SheetModel()
            {
                SpreadsheetId = spreadsheetId,
                Title = sheetName
            };

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, sheetName);
            IList<IList<object>> data = GetData(spreadsheetId, sheetName);

            sheetModel.SpreadsheetTitle = spreadsheet.Properties.Title;
            sheetModel.Gid = sheet.Properties.SheetId.Value;
            sheetModel.Mode = SheetMode.HeadAndKey;
            sheetModel.KeyName = keyName;

            sheetModel.ValidateData(data, keyName);
            sheetModel.Fill(data);

            return sheetModel;
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
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="OAuthSheetsScopeException"></exception>
        public void UpdateSheet(SheetModel sheetModel)
        {
            CheckSheetService();
            CheckPrincipal();

            try
            {
                CreateAppendRequest(sheetModel)?.Execute();
                CreateUpdateRequest(sheetModel)?.Execute();
                CreateDeleteRequest(sheetModel)?.Execute();

                sheetModel.ClearDeletedRows();
                sheetModel.RenumberRows();
                sheetModel.ResetRowStatuses();
            }
            catch (GoogleApiException e) when (e.HttpStatusCode == HttpStatusCode.Forbidden && e.Error.Message.Contains("insufficient authentication scopes"))
            {
                throw new OAuthSheetsScopeException(e.Error.Message, e);
            }
            catch (GoogleApiException e) when (e.HttpStatusCode == HttpStatusCode.Forbidden && e.Error.Message.Contains("does not have permission"))
            {
                throw new UserAccessDeniedException(e.Error.Message, e)
                {
                    Operation = $"Обновление листа: {sheetModel.SpreadsheetTitle}/{sheetModel.Title}",
                };
            }
        }


        #region CheckFields
        /// <exception cref="InvalidOperationException"></exception>
        private void CheckSheetService()
        {
            if (_sheetsService == null)
            {
                throw new InvalidOperationException(
                    $"Необходимо аутентифицироваться перед использованием класса {nameof(GCPApplication)}."
                );
            }
        }

        /// <exception cref="UserAccessDeniedException"></exception>
        private void CheckPrincipal()
        {
            if (_principal is ServiceAccount)
            {
                throw new UserAccessDeniedException(
                    $"Нельзя изменять листы googl таблиц, используя {nameof(ServiceAccount)}.")
                {
                    Operation = "Обновление листа"
                };
            }
        }
        #endregion

        #region Spreadsheets
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        private Spreadsheet GetGoogleSpreadsheet(string spreadsheetId)
        {
            try
            {
                return _sheetsService.Spreadsheets.Get(spreadsheetId).Execute();
            }
            catch (GoogleApiException e) when (e.HttpStatusCode == HttpStatusCode.BadRequest)
            {
                // Спорное и не явное решение,
                // но пока что данное состояние было только из-за невалидного api key.
                // Возможно следует найти уточняющее состояние ошибки.
                throw new InvalidApiKeyException($"Не удалось получить таблицу с id: {spreadsheetId}", e);
            }
            catch (GoogleApiException e) when (e.HttpStatusCode == HttpStatusCode.Forbidden && e.Error.Message.Contains("does not have permission"))
            {
                throw new UserAccessDeniedException(e.Error.Message, e)
                {
                    Operation = $"Получение таблицы с id: {spreadsheetId}"
                };
            }
            catch (GoogleApiException e) when (e.HttpStatusCode == HttpStatusCode.NotFound)
            {
                throw new SpreadsheetNotFoundException(e.Error.Message, e)
                {
                    SpreadsheetId = spreadsheetId
                };
            }
        }

        /// <exception cref="SheetNotFoundException"></exception>
        private Sheet GetGoogleSheet(Spreadsheet spreadsheet, int gid)
        {
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.SheetId == gid);

            if (sheet == null)
            {
                throw new SheetNotFoundException(
                    $"В таблице {spreadsheet.Properties.Title} " +
                    $"с id: {spreadsheet.SpreadsheetId} " +
                    $"нет листа с gid: {gid}"
                )
                {
                    SpreadsheetName = spreadsheet.Properties.Title,
                    SpreadsheetId = spreadsheet.SpreadsheetId,
                    SheetGid = gid.ToString()
                };
            }

            return sheet;
        }

        /// <exception cref="SheetNotFoundException"></exception>
        private Sheet GetGoogleSheet(Spreadsheet spreadsheet, string sheetName)
        {
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetName);

            if (sheet == null)
            {
                throw new SheetNotFoundException(
                    $"В таблице {spreadsheet.Properties.Title} " +
                    $"с id: {spreadsheet.SpreadsheetId} " +
                    $"нет листа с именем: {sheetName}"
                )
                {
                    SpreadsheetName = spreadsheet.Properties.Title,
                    SpreadsheetId = spreadsheet.SpreadsheetId,
                    SheetName = sheetName
                };
            }

            return sheet;
        }

        private IList<IList<object>> GetData(string spreadsheetId, string sheetName)
        {
            return _sheetsService
                .Spreadsheets
                .Values
                .Get(spreadsheetId, sheetName)
                .Execute()
                .Values ?? new List<IList<object>>();
        }
        #endregion

        #region UpdateSheetModel
        private SpreadsheetsResource.ValuesResource.BatchUpdateRequest CreateUpdateRequest(SheetModel sheet)
        {
            if (sheet.Rows.FindAll(row => row.Status == RowStatus.ToChange).Count <= 0)
            {
                return null;
            }
 
            var requestBody = new BatchUpdateValuesRequest
            {
                Data = sheet.GetChangeValueRange(),
                ValueInputOption = SpreadsheetsResource
                    .ValuesResource
                    .AppendRequest
                    .ValueInputOptionEnum
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
            if (sheet.Rows.FindAll(row => row.Status == RowStatus.ToDelete).Count <= 0)
            {
                return null;
            }

            var requestBody = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>()
            };

            foreach (List<Row> groupRows in sheet.GetDeleteRows())
            {
                requestBody.Requests.Add(
                    CreateDeleteDimensionRequest(
                        sheet.Gid,
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
            if (sheet.Rows.FindAll(row => row.Status == RowStatus.ToAppend).Count <= 0)
            {
                return null;
            }
 
            var request = _sheetsService
                .Spreadsheets
                .Values
                .Append(sheet.GetAppendValueRange(), sheet.SpreadsheetId, sheet.Title);

            request.ValueInputOption = SpreadsheetsResource
                .ValuesResource
                .AppendRequest
                .ValueInputOptionEnum
                .USERENTERED;

            return request;
        }
        #endregion
    }
}
