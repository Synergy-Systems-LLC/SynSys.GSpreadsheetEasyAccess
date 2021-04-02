using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace GetGoogleSheetDataAPI
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
        /// Попытка подключения к приложению на Google Cloud Platform.<br/>
        /// Для пользователя это выглядит как предложение к подключению в браузере.
        /// <example>
        /// <code>
        /// bool isConnectionSuccessful = connector.TryToCreateConnect(new MemoryStream(Properties.Resources.credentials));
        /// </code>
        /// </example>
        /// </summary>
        /// <remarks>
        /// Если пользователь не входит в домент synsys.co, то ему будет отказано в подлючении.<br/>
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

        #region GetData

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet.<br/>
        /// Лист представляет таблицу из списка строк. Шапка отсутствует.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение.
        /// </summary>
        /// <param name="url">Полный url адрес листа</param>
        /// <param name="sheet"></param>
        /// <returns>
        /// Всегда вернёт объект типа Sheet из out параметра, но если возникнет ошибка,
        /// то в Sheet.Status будет её значение, а в Sheet.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSimpleSheet(string url, out Sheet sheet)
        {
            sheet = new Sheet();
 
            try
            {
                InitializeSheet(sheet, url, SheetMode.Simple, string.Empty);
 
                var data = GetData(sheet);

                if (data != null)
                {
                    sheet.Fill(data);
                }
            }
            catch (Exception exception)
            {
                sheet.Status = exception.ToString();
                return false;
            }

            return true;
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet.<br/>
        /// Лист представляет таблицу из списка строк, без первой строки.
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="url">Полный url адрес листа</param>
        /// <param name="sheet"></param>
        /// <returns>
        /// Всегда вернёт объект типа Sheet из out параметра, но если возникнет ошибка,
        /// то в Sheet.Status будет её значение, а в Sheet.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSheetWithHead(string url, out Sheet sheet)
        {
            sheet = new Sheet();

            try
            {
                InitializeSheet(sheet, url, SheetMode.Head, string.Empty);
 
                var data = GetData(sheet);

                if (data == null)
                {
                    sheet.Status = $"Лист по адресу {url} не содержит данных";
                    return false;
                }

                sheet.Fill(data);
            }
            catch (Exception exception)
            {
                sheet.Status = exception.ToString();
                return false;
            }

            return true;
        }

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet.<br/>
        /// Лист представляет таблицу из списка строк, без первой строки.
        /// Первая строка уходит в шапку таблицы.<br/>
        /// В каждой строке одинаковое количество ячеек и есть ключевая ячейка.<br/>
        /// В каждой ячейке есть строковое значение и наименование,
        /// которое совпадает с шапкой столбца для данной ячейки.
        /// </summary>
        /// <param name="url">Полный url адрес листа</param>
        /// <param name="keyName">Наименование ключевого столбца таблицы</param>
        /// <param name="sheet"></param>
        /// <returns>
        /// Всегда вернёт объект типа Sheet из out параметра, но если возникнет ошибка,
        /// то в Sheet.Status будет её значение, а в Sheet.Rows будет пустой список.
        /// </returns>
        public bool TryToGetSheetWithHeadAndKey(string url, string keyName, out Sheet sheet)
        {
            sheet = new Sheet();
 
            try
            {
                InitializeSheet(sheet, url, SheetMode.HeadAndKey, keyName);
 
                var data = GetData(sheet);

                if (data == null)
                {
                    sheet.Status = $"Лист по адресу {url} не содержит данных";
                    return false;
                }

                if (!data[0].Contains(keyName))
                {
                    sheet.Status = $"Шапка листа по адресу {url} не содержит столбец \"{keyName}\"";
                    return false;
                }

                sheet.Fill(data);
            }
            catch (Exception exception)
            {
                sheet.Status = exception.ToString();
                return false;
            }

            return true;
        }

        private void InitializeSheet(Sheet sheet, string url, SheetMode mode, string keyName)
        {
            sheet.Mode = mode;
            sheet.KeyName = keyName;

            if (!HttpManager.IsCorrectUrl(url))
            {
                throw new Exception(HttpManager.Status);
            }

            sheet.SpreadsheetId = HttpManager.GetSpreadsheetIdFromUrl(url);
            sheet.Gid = HttpManager.GetGidFromUrl(url);
            sheet.Title = GetSheetTitle(sheet);
            sheet.SpreadsheetTitle = GetSpreadsheetTitle(sheet);
        }

        private string GetSpreadsheetTitle(Sheet sheet)
        {
            return GetSpreadsheet(sheet.SpreadsheetId).Properties.Title;
        }

        /// <summary>
        /// Метод может бросить исключение System.NullReferenceException
        /// если по какой-то причине sheetService не будет получен.<br/>
        /// Метод может бросить исключение Google.Apis.Requests.RequestError в ValueResource.Get
        /// если id или gid листа не существуют.
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private IList<IList<object>> GetData(Sheet sheet)
        {
            return sheetsService
                .Spreadsheets
                .Values
                .Get(sheet.SpreadsheetId, sheet.Title)
                .Execute()
                .Values;
        }

        /// <summary>
        /// Метод может бросить исключение Google.Apis.Requests.RequestError
        /// в том случае если id или gid листа не существуют.
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private string GetSheetTitle(Sheet sheet)
        {
            return (from _sheet in GetSpreadsheet(sheet.SpreadsheetId).Sheets
                    where _sheet.Properties.SheetId.ToString() == sheet.Gid
                    select _sheet.Properties.Title).FirstOrDefault();
        }

        /// <summary>
        /// Метод может бросить исключение Google.Apis.Requests.RequestError
        /// в том случае если id таблицы не существуют.
        /// </summary>
        /// <param name="spreadsheetId">Число из url листа. /Id/</param>
        /// <returns></returns>
        private Spreadsheet GetSpreadsheet(string spreadsheetId)
        {
            return sheetsService
                .Spreadsheets
                .Get(spreadsheetId)
                .Execute();
        }

        /// <summary>
        /// Обновленние листа google таблицы на основе изменённого экземпляра типа Sheet.
        /// <remarks>
        /// Метод изменяет данные в ячейках,
        /// добавляет строки в конец листа и удаляет выбраные строки.
        /// Все эти действия просходят на основе запросов в google.
        /// </remarks>
        /// </summary>
        /// <param name="sheet"></param>
        public void UpdateSheet(Sheet sheet)
        {
            CreateAppendRequest(sheet)?.Execute();
            CreateUpdateRequest(sheet)?.Execute();
            CreateDeleteRequest(sheet)?.Execute();

            sheet.ClearDeletedRows();
            sheet.RenumberRows();
            sheet.ResetRowStatuses();
        }
        #endregion

        #region Requests
        /// <summary>
        /// Создание запроса на обновление листа в google таблице.
        /// В запросе содержатся изменения значений в ячейках листа.
        /// </summary>
        /// <param name="sheet">Экземпляр изменённого листа</param>
        /// <returns></returns>
        private SpreadsheetsResource.ValuesResource.BatchUpdateRequest CreateUpdateRequest(Sheet sheet)
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
        private SpreadsheetsResource.BatchUpdateRequest CreateDeleteRequest(Sheet sheet)
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
        private SpreadsheetsResource.ValuesResource.AppendRequest CreateAppendRequest(Sheet sheet)
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
    }
}
