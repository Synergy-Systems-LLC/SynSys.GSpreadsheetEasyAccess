using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace SynSys.GSpreadsheetEasyAccess
{
    /// <summary>
    /// Статусы подключения к Google Cloud Platform.
    /// </summary>
    public enum ConnectStatus
    {
        /// <summary>
        /// Кoннектор подключён к приложению на Google Cloud Platform.
        /// </summary>
        Connected,
        /// <summary>
        /// Коннектор не подключён к приложению на Google Cloud Platform.
        /// </summary>
        NotConnected,
        /// <summary>
        /// Коннектор не подключён к приложению на Google Cloud Platform
        /// из-за того что вышло время на подлючение.
        /// </summary>
        AuthorizationTimedOut
    }

    /// <summary>
    /// Connector - это обёртка для библиотек Google.Apis.<br/>
    /// Через него происходит подключение к приложению на Google Cloup Platform,
    /// к которому в свою очередь подключён вэб сервис Google Sheets API.<br/>
    /// Подключение проиходит с помощью получения файла credentials.json данного приложения.
    /// </summary>
    /// <remarks>
    /// После успешного подключения можно получать листы гугл таблиц по их id и gid.
    /// Коннектор так же позволяет обновлять данные в листах.
    /// </remarks>
    public class Connector
    {
        private readonly string[] scopes = { SheetsService.Scope.Spreadsheets };
        private SheetsService sheetsService;
 
        /// <summary>
        /// Имя приложения, которое будет использоваться в заголовке User-Agent.<br/>
        /// Значение по умолчанию <c>string.Empty</c>. 
        /// </summary>
        public string ApplicationName { get; set; } = string.Empty;
        /// <summary>
        /// Время на попытку подключения к приложению на Google Cloud Platform в секундах.<br/>
        /// Значение по умолчанию <c>15</c>. 
        /// </summary>
        public byte CancellationSeconds { get; set; } = 15;
        /// <summary>
        /// В зависимости от состояния подключения меняется его статус.<br/>
        /// Значение по умолчанию <c>NotConnected</c>
        /// </summary>
        /// <value>
        /// Статус текущего подключения.
        /// </value>
        public ConnectStatus Status { get; private set; } = ConnectStatus.NotConnected;
        /// <summary>
        /// Исключение возникшее во время подключения к приложению на Google Cloud Platform.
        /// Значение по умолчанию <c>new Exception()</c>.
        /// </summary>
        public Exception Exception { get; private set; } = new Exception();

        /// <summary>
        /// Попытка подключения к приложению на Google Cloud Platform, а именно к сервису таблиц.<br/>
        /// Для пользователя это выглядит как предложение аутентификации в браузере.
        /// <example>
        /// <code>
        /// bool isConnectionSuccessful = connector.TryToCreateConnect(new MemoryStream(Properties.Resources.credentials));
        /// </code>
        /// </example>
        /// </summary>
        /// <remarks>
        /// У пользователя есть время на подключение, оно настраивается через свойство CancellationSeconds.<br/>
        /// После истечения данного времени коннектору присваивается статус ConnectorStatus.AuthorizationTimeOut.<br/>
        /// Если подключение состоялось, то присваивается статус ConnectorStatus.Connected 
        /// и коннектором можно пользоваться для получения и изменения листов таблицы.
        /// </remarks>
        /// <param name="credentials">Экземпляр класса Stream полученный из файла credentials.json</param>
        public bool TryToCreateConnect(Stream credentials) 
        {
            try
            {
                sheetsService = GetService(credentials);
                Status = ConnectStatus.Connected;
            }
            catch (AggregateException ax)
            {
                Exception = ax;
                Status = ConnectStatus.NotConnected;

                foreach (var e in ax.InnerExceptions)
                {
                    if (e is TaskCanceledException)
                    {
                        Exception = e;
                        Status = ConnectStatus.AuthorizationTimedOut;
                    }
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представляет таблицу из списка строк. Шапка отсутствует.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение.
        /// </summary>
        /// <param name="uri">Полный uri листа гугл таблицы</param>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <returns>
        /// Всегда вернёт объект типа SheetModel из out параметра.<br/>
        /// Если при получении данных возникнет ошибка, то в sheetModel.Status будет её значение,
        /// а в sheetModel.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSimpleSheet(string uri, out SheetModel sheetModel)
        {
            return TryToGetSimpleSheet(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri),
                out sheetModel
            );
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представляет таблицу из списка строк. Шапка отсутствует.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение.
        /// </summary>
        /// <param name="spreadsheetId">Идентификатор гугл таблицы</param>
        /// <param name="gid">Идентификатор листа</param>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <returns>
        /// Всегда вернёт объект типа SheetModel из out параметра.<br/>
        /// Если получении данных возникнет ошибка, то в sheetModel.Status будет её значение,
        /// а в SheetModel.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSimpleSheet(string spreadsheetId,
                                        int gid,
                                        out SheetModel sheetModel)
        {
            bool isInitializeSuccessful = TryToInitializeSheet(
                spreadsheetId,
                gid,
                SheetMode.Simple,
                string.Empty,
                out sheetModel
            );

            return TryToFillSimpleSheet(sheetModel, isInitializeSuccessful);
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представляет таблицу из списка строк. Шапка отсутствует.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение.
        /// </summary>
        /// <param name="spreadsheetId">Идентификатор гугл таблицы</param>
        /// <param name="sheetName">Имя листа</param>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <returns>
        /// Всегда вернёт объект типа SheetModel из out параметра.<br/>
        /// Если получении данных возникнет ошибка, то в sheetModel.Status будет её значение,
        /// а в SheetModel.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSimpleSheet(string spreadsheetId,
                                        string sheetName,
                                        out SheetModel sheetModel)
        {
            bool isInitializeSuccessful = TryToInitializeSheet(
                spreadsheetId,
                sheetName,
                SheetMode.Simple,
                string.Empty,
                out sheetModel
            );

            return TryToFillSimpleSheet(sheetModel, isInitializeSuccessful);
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа SheetModel.<br/>
        /// Лист представляет таблицу из списка строк, без первой строки.
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="uri">Полный uri адрес листа</param>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <returns>
        /// Всегда вернёт объект типа SheetModel из out параметра.<br/>
        /// Если получении данных возникнет ошибка, то в sheetModel.Status будет её значение,
        /// а в SheetModel.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSheetWithHead(string uri, out SheetModel sheetModel)
        {
            return TryToGetSheetWithHead(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri),
                out sheetModel
            );
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet.<br/>
        /// Лист представляет таблицу из списка строк, без первой строки.
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="spreadsheetId">Идентификатор гугл таблицы</param>
        /// <param name="gid">Идентификатор листа</param>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <returns>
        /// Всегда вернёт объект типа SheetModel из out параметра.<br/>
        /// Если получении данных возникнет ошибка, то в sheetModel.Status будет её значение,
        /// а в SheetModel.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSheetWithHead(string spreadsheetId,
                                          int gid,
                                          out SheetModel sheetModel)
        {
            bool isInitializeSuccessful = TryToInitializeSheet(
                spreadsheetId,
                gid,
                SheetMode.Head,
                string.Empty,
                out sheetModel
            );

            return TryToFillSimpleSheet(sheetModel, isInitializeSuccessful);
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet.<br/>
        /// Лист представляет таблицу из списка строк, без первой строки.
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="spreadsheetId">Идентификатор гугл таблицы</param>
        /// <param name="sheetName">Имя листа</param>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <returns>
        /// Всегда вернёт объект типа SheetModel из out параметра.<br/>
        /// Если получении данных возникнет ошибка, то в sheetModel.Status будет её значение,
        /// а в SheetModel.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSheetWithHead(string spreadsheetId,
                                          string sheetName,
                                          out SheetModel sheetModel)
        {
            bool isInitializeSuccessful = TryToInitializeSheet(
                spreadsheetId,
                sheetName,
                SheetMode.Head,
                string.Empty,
                out sheetModel
            );

            return TryToFillSimpleSheet(sheetModel, isInitializeSuccessful);
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet.<br/>
        /// Лист представляет таблицу из списка строк, без первой строки.
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек и есть ключевая ячейка.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="uri">Полный uri адрес листа</param>
        /// <param name="keyName">Имя ключевого столбца.</param>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <returns>
        /// Всегда вернёт объект типа SheetModel из out параметра.<br/>
        /// Если получении данных возникнет ошибка, то в sheetModel.Status будет её значение,
        /// а в SheetModel.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSheetWithHeadAndKey(string uri,
                                                string keyName,
                                                out SheetModel sheetModel)
        {
            return TryToGetSheetWithHeadAndKey(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri),
                keyName,
                out sheetModel
            );
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet.<br/>
        /// Лист представляет таблицу из списка строк, без первой строки.
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек и есть ключевая ячейка.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="spreadsheetId">Идентификатор гугл таблицы</param>
        /// <param name="gid">Идентификатор листа</param>
        /// <param name="keyName">Имя ключевого столбца.</param>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <returns>
        /// Всегда вернёт объект типа SheetModel из out параметра.<br/>
        /// Если получении данных возникнет ошибка, то в sheetModel.Status будет её значение,
        /// а в SheetModel.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSheetWithHeadAndKey(string spreadsheetId,
                                                int gid,
                                                string keyName,
                                                out SheetModel sheetModel)
        {
            bool isInitializeSuccessful = TryToInitializeSheet(
                spreadsheetId,
                gid,
                SheetMode.HeadAndKey,
                keyName,
                out sheetModel
            );

            return TryToFillSheetWithHeadAndKey(sheetModel, keyName, isInitializeSuccessful);
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet.<br/>
        /// Лист представляет таблицу из списка строк, без первой строки.
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек и есть ключевая ячейка.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="spreadsheetId">Идентификатор гугл таблицы</param>
        /// <param name="sheetName">Имя листа</param>
        /// <param name="keyName">Имя ключевого столбца.</param>
        /// <param name="sheetModel">Модель листа гугл таблицы</param>
        /// <returns>
        /// Всегда вернёт объект типа SheetModel из out параметра.<br/>
        /// Если получении данных возникнет ошибка, то в sheetModel.Status будет её значение,
        /// а в SheetModel.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSheetWithHeadAndKey(string spreadsheetId,
                                                string sheetName,
                                                string keyName,
                                                out SheetModel sheetModel)
        {
            bool isInitializeSuccessful = TryToInitializeSheet(
                spreadsheetId,
                sheetName,
                SheetMode.HeadAndKey,
                keyName,
                out sheetModel
            );

            return TryToFillSheetWithHeadAndKey(sheetModel, keyName, isInitializeSuccessful);
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
        public void UpdateSheet(SheetModel sheetModel)
        {
            CreateAppendRequest(sheetModel)?.Execute();
            CreateUpdateRequest(sheetModel)?.Execute();
            CreateDeleteRequest(sheetModel)?.Execute();

            sheetModel.ClearDeletedRows();
            sheetModel.RenumberRows();
            sheetModel.ResetRowStatuses();
        }


        #region WorkWithSheetModel
        private bool TryToInitializeSheet(string spreadsheetId,
                                          int gid,
                                          SheetMode mode,
                                          string keyName,
                                          out SheetModel sheetModel)
        {
            sheetModel = new SheetModel();
            var commonMessage = "Во время инициализации листа";

            if (string.IsNullOrWhiteSpace(spreadsheetId))
            {
                sheetModel.Status = "HttpUtils не смог распознать spreadsheetId в ведённом uri";
                return false;
            }

            if (gid < 0)
            {
                sheetModel.Status = $"Был введён некорректный gid листа: {gid}. Возможно он отсутствует в uri";
                return false;
            }

            if (!TryToGetSpreadsheet(spreadsheetId, out Spreadsheet spreadsheet, out string exceptionMessage))
            {
                sheetModel.Status = $"{commonMessage} {exceptionMessage}";
                return false;
            }

            if (SheetNotFound(spreadsheet.Sheets, gid, out Sheet sheet))
            {
                sheetModel.Status = $"{commonMessage} " +
                                    $"в таблице Id: \"{spreadsheet.SpreadsheetId}\" " +
                                    $"не найден лист Id: {gid}";
                return false;
            }

            sheetModel.Mode = mode;
            sheetModel.KeyName = keyName;
            sheetModel.Gid = gid.ToString();
            sheetModel.Title = sheet.Properties.Title;
            sheetModel.SpreadsheetId = spreadsheet.SpreadsheetId;
            sheetModel.SpreadsheetTitle = spreadsheet.Properties.Title;

            return true;
        }

        private bool TryToInitializeSheet(string spreadsheetId,
                                          string sheetName,
                                          SheetMode mode,
                                          string keyName,
                                          out SheetModel sheetModel)
        {
            sheetModel = new SheetModel();
            var commonMessage = "Во время инициализации листа";

            if (string.IsNullOrWhiteSpace(spreadsheetId))
            {
                sheetModel.Status = "HttpUtils не смог распознать spreadsheetId в ведённом uri";
                return false;
            }

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                sheetModel.Status = $"Былo введёнo пустое имя листа";
                return false;
            }

            if (!TryToGetSpreadsheet(spreadsheetId, out Spreadsheet spreadsheet, out string exceptionMessage))
            {
                sheetModel.Status = $"{commonMessage} {exceptionMessage}";
                return false;
            }

            if (SheetNotFound(spreadsheet.Sheets, sheetName, out Sheet sheet))
            {
                sheetModel.Status = $"{commonMessage} " +
                                    $"в таблице Id: \"{spreadsheet.SpreadsheetId}\" " +
                                    $"не найден лист с именем: {sheetName}";
                return false;
            }

            sheetModel.Mode = mode;
            sheetModel.KeyName = keyName;
            sheetModel.Gid = sheet.Properties.SheetId.ToString();
            sheetModel.Title = sheet.Properties.Title;
            sheetModel.SpreadsheetId = spreadsheet.SpreadsheetId;
            sheetModel.SpreadsheetTitle = spreadsheet.Properties.Title;

            return true;
        }

        private bool TryToGetSpreadsheet(string spreadsheetId,
                                         out Spreadsheet spreadsheet,
                                         out string exceptionMessage)
        {
            spreadsheet = null;
            exceptionMessage = string.Empty;

            try
            {
                spreadsheet = sheetsService.Spreadsheets.Get(spreadsheetId).Execute();
                return true;
            }
            catch (GoogleApiException e)
            {
                if (e.Error == null)
                {
                    exceptionMessage = $"служба Google API выбросила исключение: {e.Message}";
                }
                else
                {
                    exceptionMessage = $"служба Google API выбросила исключение: {e.Error.Message}\n" +
                                       $"{HttpUtils.GetMessageByCode(e.Error.Code.ToString())}";
                }

                return false;
            }
            catch (TokenResponseException)
            {
                exceptionMessage = $"возникла пробема с токеном доступа. " +
                                   $"Старый токен удалён, а при следующем запуске будет создан новый.";
                return false;
            }
            catch (Exception e)
            {
                exceptionMessage = $"возникла непредвиденная ошибка: {e}";
                return false;
            }
        }

        private bool SheetNotFound(IList<Sheet> sheets, string sheetName, out Sheet sheet)
        {
            sheet = (from _sheet in sheets
                     where _sheet.Properties.Title == sheetName
                     select _sheet).FirstOrDefault();

            return sheet == null;
        }

        private bool SheetNotFound(IList<Sheet> sheets, int gid, out Sheet sheet)
        {
            sheet = (from _sheet in sheets
                     where _sheet.Properties.SheetId == gid
                     select _sheet).FirstOrDefault();

            return sheet == null;
        }

        private bool TryToFillSimpleSheet(SheetModel sheetModel, bool isInitializeSuccessful)
        {
            if (!isInitializeSuccessful)
            {
                return false;
            }

            IList<IList<object>> data;

            try
            {
                data = GetData(sheetModel);
            }
            catch(Exception e)
            {
                sheetModel.Status = $"Во время получения данных возникло исключение: {e}";
                return false;
            }

            if (data == null)
            {
                sheetModel.Status = $"Лист не содержит данных";
                return false;
            }

            sheetModel.Fill(data);

            return true;
        }

        private bool TryToFillSheetWithHeadAndKey(SheetModel sheetModel, string keyName, bool isInitializeSuccessful)
        {
            if (!isInitializeSuccessful)
            {
                return false;
            }

            IList<IList<object>> data;

            try
            {
                data = GetData(sheetModel);
            }
            catch(Exception e)
            {
                sheetModel.Status = $"Во время получения данных возникло исключение: {e}";
                return false;
            }

            if (data == null)
            {
                sheetModel.Status = $"Лист не содержит данных";
                return false;
            }

            if (!data[0].Contains(keyName))
            {
                sheetModel.Status = $"Лист не содержит ключ {keyName}";
                return false;
            }

            sheetModel.Fill(data);

            return true;
        }

        /// <summary>
        /// Метод может бросить исключение System.NullReferenceException
        /// если по какой-то причине sheetService не будет получен.<br/>
        /// Метод может бросить исключение Google.Apis.Requests.RequestError в ValueResource.Get
        /// если id или gid листа не существуют.
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private IList<IList<object>> GetData(SheetModel sheet)
        {
            return sheetsService
                .Spreadsheets
                .Values
                .Get(sheet.SpreadsheetId, sheet.Title)
                .Execute()
                .Values;
        }
        #endregion

        #region Requests
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

            return sheetsService
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
                    sheetsService,
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
 
            var request = sheetsService
                .Spreadsheets
                .Values
                .Append(sheet.GetAppendValueRange(), sheet.SpreadsheetId, $"{sheet.Title}!A:A");

            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum
                .USERENTERED;

            return request;
        }
        #endregion

        #region SheetService
        private SheetsService GetService(Stream credentials)
        {
            return new SheetsService(
                new BaseClientService.Initializer
                {
                    HttpClientInitializer = GetUserCredential(credentials),
                    ApplicationName = ApplicationName
                }
            );
        }

        private UserCredential GetUserCredential(Stream credentials)
        {
            return GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(credentials).Secrets,
                scopes,
                "user",
                GenerateCancellationToken(),
                new FileDataStore(GenerateTokenPath(), true)
            ).Result;
        }

        private CancellationToken GenerateCancellationToken()
        {
            var tokenSource = new CancellationTokenSource();
            tokenSource.CancelAfter(TimeSpan.FromSeconds(CancellationSeconds));
            return tokenSource.Token;
        }

        private static string GenerateTokenPath()
        {
            var uriPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().EscapedCodeBase);
            var outputDirectory = new Uri(uriPath).LocalPath;
            return Path.Combine(outputDirectory, "token.json");
        }
        #endregion
    }
}
