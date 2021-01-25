using System;
using System.IO;
using System.Threading;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using System.Threading.Tasks;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;

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
    /// Connector - это обёртка для библиотек Google.Apis.
    /// Через него происходит подключение к приложению "get google sheet data", которое находится на Google Cloup Platform.
    /// Подключение проиходит с помощью получения файла credentials.json данного приложения (get google sheet data).
    /// </summary>
    /// <remarks>
    /// После успешного подключения можно получать листы гугл таблиц по их id и gid.
    /// Коннектор так же позволяет обновлять данные в листах.
    /// </remarks>
    public class Connector
    {
        string[] scopes = { SheetsService.Scope.Spreadsheets };
        string applicationName = "Get Google Sheet Data";
        byte cancellationSeconds = 15;
        SheetsService sheetsService;

        /// <summary>
        /// В зависимости от состояния подключения меняется его статус.
        /// </summary>
        /// <value>
        /// Статус текущего подключения.
        /// </value>
        public ConnectStatus Status { get; private set; } = ConnectStatus.NotConnected;

        /// <summary>
        /// Инициализирует экземпляр коннектора.
        /// Данный коннектор не имеет подключения к приложению на Google Cloud Platform.
        /// Чтобы создать подключение необходимо вызвать метод TryToCreateCo
        /// </summary>
        public Connector() { }

        /// <summary>
        /// Попытка подключения к приложению на Google Cloud Platform.
        /// Для пользователя это выглядит как предложение к подключению в браузере.
        /// <example>
        /// <code>
        /// var connector = new Connector();
        /// bool isConnectionSuccessful = connector.TryToCreateConnect(new MemoryStream(Properties.Resources.credentials));
        /// </code>
        /// </example>
        /// </summary>
        /// <remarks>
        /// У пользователя есть 15 секунд на подключение, после этого коннектору присваивается статус ConnectorStatus.AuthorizationTimeOut.
        /// Если пользователь не входит в домент synsys.co, то ему будет отказано в подлючении.
        /// В этом случае коннектор отключится через 15 секунд.
        /// Если подключене состоялось, то присваивается статус ConnectorStatus.Connected
        /// и у коннектора можно вызывать методы для получения листов и их данных.
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
                Status = ConnectStatus.NotConnected;

                if (ax.InnerExceptions == null)
                {
                    return false;
                }

                foreach (var e in ax.InnerExceptions)
                {
                    if (e is TaskCanceledException)
                    {
                        Status = ConnectStatus.AuthorizationTimedOut;
                        return false;
                    }
                }
            }

            return true;
        }

        #region SheetService
        private SheetsService GetService(Stream credentials)
        {
            return new SheetsService(
                new BaseClientService.Initializer()
                {
                    HttpClientInitializer = GetUserCredential(credentials),
                    ApplicationName = applicationName,
                }
            );
        }

        private UserCredential GetUserCredential(Stream credentials)
        {
            return GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(credentials).Secrets,
                scopes,
                "user",
                GenerateCansellationToken(),
                new FileDataStore(GenerateTokenPath(), true)
            ).Result;
        }

        private CancellationToken GenerateCansellationToken()
        {
            var tokenSource = new CancellationTokenSource();
            tokenSource.CancelAfter(TimeSpan.FromSeconds(cancellationSeconds));
            return tokenSource.Token;
        }

        private string GenerateTokenPath()
        {
            string uriPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().EscapedCodeBase);
            string outPutDirectory = new Uri(uriPath).LocalPath;
            return Path.Combine(outPutDirectory, "token.json");
        }
        #endregion

        #region GetData
        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet
        /// </summary>
        /// <param name="url">Полный url адрес листа</param>
        /// <returns>
        /// Всегда вернёт объект типа Sheet, но если url не относится к какому-то конкретному листу
        /// гугл таблицы то данный объект будет пуст и иметь статус того, почему он пуст.
        /// </returns>
        public bool TryToGetSheet(string url, SheetMode mode, out Sheet sheet)
        {
            sheet = new Sheet()
            {
                Mode = mode,
            };

            if (!HttpManager.IsCorrectUrl(url))
            {
                sheet.Status = HttpManager.Status;
                return false;
            }

            sheet.SpreadsheetId = HttpManager.GetSpreadsheetIdFromUrl(url);
            sheet.Gid = HttpManager.GetGidFromUrl(url);

            try
            {
                sheet.Title = GetSheetTitle(sheet.SpreadsheetId, sheet.Gid);
                sheet.SpreadsheetTitle = GetSpreadsheetTitle(sheet.SpreadsheetId);
                sheet.Fill(GetData(sheet.SpreadsheetId, sheet.Gid));
            }
            catch (Exception exeption)
            {
                sheet.Status = exeption.Message;
                return false;
            }

            return true;
        }

        private string GetSpreadsheetTitle(string spreadsheetId)
        {
            return GetSpreadsheet(spreadsheetId).Properties.Title;
        }

        /// <summary>
        /// Получение данных из листа гугл таблицы.
        /// </summary>
        /// <remarks>
        /// Метод можен выбросить ошибку если sheetsService is null.
        /// Так же ошибка может возникнуть в методе Get если не будет найдено имя листа.
        /// </remarks>
        /// <param name="spreadsheetId">Число из url листа. /Id/</param>
        /// <param name="gid">Число из url листа. git=Id</param>
        /// <returns></returns>
        private IList<IList<object>> GetData(string spreadsheetId, string gid)
        {
            return sheetsService
                .Spreadsheets
                .Values
                .Get(spreadsheetId, GetSheetTitle(spreadsheetId, gid))
                .Execute()
                .Values;
        }

        /// <summary>
        /// Получение имени листа по его spreadsheetId и gid.
        /// </summary>
        /// <param name="spreadsheetId">Число из url листа. /Id/</param>
        /// <param name="gid">Число из url листа. git=Id</param>
        /// <returns></returns>
        private string GetSheetTitle(string spreadsheetId, string gid)
        {
            foreach (Google.Apis.Sheets.v4.Data.Sheet sheet in GetSpreadsheet(spreadsheetId).Sheets)
            {
                if (sheet.Properties.SheetId.ToString() == gid)
                {
                    return sheet.Properties.Title;
                }
            }

            return null;
        }

        /// <summary>
        /// Получение google таблицы по её id.
        /// Получение происходит по запросу из метода Get.
        /// </summary>
        /// <remarks>
        /// Метод может выбросить ошибку если spreadsheetId не существует среди таблиц.
        /// </remarks>
        /// <param name="spreadsheetId">Число из url листа. /Id/</param>
        /// <returns></returns>
        private Spreadsheet GetSpreadsheet(string spreadsheetId)
        {
            return sheetsService.Spreadsheets.Get(spreadsheetId).Execute();
        }

        /// <summary>
        /// Обновленние листа google таблицы на основе изменённого экземпляра
        /// типа Sheet, полученного из метода TryToGetSheet.
        /// <remarks>
        /// Метод изменяет данные в ячейках,
        /// добавляет строки в конец листа и удаляет выбраные строки.
        /// Всё это просходит на основе запросов в google.
        /// </remarks>
        /// </summary>
        /// <param name="sheet"></param>
        public void UpdateSheet(Sheet sheet)
        {
            var appendRequest = CreateAppendRequest(sheet);
            var appendResponse = appendRequest?.Execute();

            var updateRequest = CreateUpdateRequest(sheet);
            var updateResponse = updateRequest?.Execute();

            var deleteRequest = CreateDeleteRequest(sheet);
            var delereResponse = deleteRequest?.Execute();

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
            if (sheet.Rows.FindAll(row => row.Status == RowStatus.ToChange).Count > 0)
            {
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

                return sheetsService
                    .Spreadsheets
                    .Values
                    .BatchUpdate(requestBody, sheet.SpreadsheetId);
            }

            return null;
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
            return new Request()
            {
                DeleteDimension = new DeleteDimensionRequest()
                {
                    Range = new DimensionRange()
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
            if (sheet.Rows.FindAll(row => row.Status == RowStatus.ToAppend).Count > 0)
            {
                var request = sheetsService
                    .Spreadsheets
                    .Values
                    .Append(sheet.GetAppendValueRange(), sheet.SpreadsheetId, $"{sheet.Title}!A:A");

                request.ValueInputOption = SpreadsheetsResource
                    .ValuesResource
                    .AppendRequest
                    .ValueInputOptionEnum
                    .USERENTERED;

                return request;
            }

            return null;
        }
        #endregion
    }
}
