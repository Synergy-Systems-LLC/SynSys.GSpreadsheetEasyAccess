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

namespace GetGoogleSheetDataAPI
{
    /// <summary>
    /// Статусы подключения к Google Cloud Platform.
    /// </summary>
    public enum ConnectStatus
    {
        /// <summary>
        /// 
        /// </summary>
        Connected,
        /// <summary>
        /// 
        /// </summary>
        NotConnected,
        /// <summary>
        /// 
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
        /// Чтобы создать подключение необходимо вызвать метод TryToCreateConnect.
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

        /// <summary>
        /// Попытка получить данные из листа гугл таблицы в виде объекта типа Sheet
        /// </summary>
        /// <param name="url">Полный url адрес листа</param>
        /// <returns>
        /// Всегда вернёт объект типа Sheet, но если url не относится к какому-то конкретному листу
        /// гугл таблицы то данный объект будет пуст и иметь статус того, почему он пуст.
        /// </returns>
        public bool TryToGetSheet(string url, out Sheet sheet)
        {
            sheet = new Sheet();

            if (!HttpManager.IsCorrectUrl(url))
            {
                sheet.Status = HttpManager.Status;
                return false;
            }

            sheet.SpreadsheetId = HttpManager.GetSpreadsheetIdFromUrl(url);
            sheet.Gid = HttpManager.GetGidFromUrl(url);

            try
            {
                sheet.Title = GetSheetName(sheet.SpreadsheetId, sheet.Gid);
                sheet.Fill(GetData(sheet.SpreadsheetId, sheet.Gid));
            }
            catch (Exception exeption)
            {
                sheet.Status = exeption.Message;
                return false;
            }

            return true;
        }

        public IList<IList<object>> GetData(string spreadsheetId, string gid)
        {
            return sheetsService
                .Spreadsheets
                .Values
                .Get(spreadsheetId, GetSheetName(spreadsheetId, gid))
                .Execute()
                .Values;
        }

        public string GetSheetName(string spreadsheetId, string gid)
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

        private Spreadsheet GetSpreadsheet(string spreadsheetId)
        {
            return sheetsService.Spreadsheets.Get(spreadsheetId).Execute();
        }
    }
}
