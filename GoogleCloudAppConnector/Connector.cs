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
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Security.Policy;

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
    /// Через него происходит подключение к приложению "get google sheet data", которое находится на Google Cloup Platform (далее GCP).
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

        /// <summary>
        /// В зависимости от состояния подключения меняется его статус.
        /// </summary>
        /// <value>
        /// Статус текущего подключения.
        /// </value>
        public ConnectStatus Status { get; private set; } = ConnectStatus.NotConnected;

        /// <summary>
        /// Если во время подключения возникла ошибка, то она попадает сюда.
        /// </summary>
        /// <value>
        /// Ошибка текущего подключения.
        /// </value>
        public Exception Exception { get; private set; }
        SheetsService sheetsService;

        private Connector() { }

        /// <summary>
        /// Создаёт новый экземпляр класса Connector, который тут же пытается подключится к приложению GCP и предлогает это сделать пользователю через браузер.
        /// <example>
        /// <code>
        /// var connector = new Connector(new MemoryStream(Properties.Resources.credentials));
        /// </code>
        /// </example>
        /// </summary>
        /// <remarks>
        /// У пользователя есть 15 секунд на подключение, после этого коннектору присваивается статус ConnectorStatus.AuthorizationTimeOut, а в свойство Exception передаётся соответсвующая ошибка.
        /// Если пользователь не входит в домент synsys.co, то ему будет отказано в подлючении.
        /// Статус подключения подключения по умолчанию ConnectStatus.NotConnected не изменится.
        /// Если подключене состоялось, то присваивается статус ConnectorStatus.Connected.
        /// </remarks>
        /// <param name="credentials">Экземпляр класса Stream полученный из файла credentials.json</param>
        public static bool TryToCreateConnect(Stream credentials, out Connector connector)
        {
            connector = new Connector();

            try
            {
                connector.GetService(credentials);
                connector.Status = ConnectStatus.Connected;
            }
            catch (AggregateException ax)
            {
                connector.Status = ConnectStatus.NotConnected;

                if (ax.InnerExceptions == null)
                {
                    return false;
                }

                foreach (var e in ax.InnerExceptions)
                {
                    if (e is TaskCanceledException)
                    {
                        connector.Status = ConnectStatus.AuthorizationTimedOut;
                        return false;
                    }
                }
            }

            return true;
        }

        #region SheetService
        private void GetService(Stream credentials)
        {
            sheetsService = new SheetsService(
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
        /// Получение данных из листа гугл таблицы в виде объекта типа Sheet
        /// </summary>
        /// <param name="url">Полный url адрес листа</param>
        /// <returns></returns>
        public Sheet TryToCreateSheet(string url)
        {
            return TryToCreateSheet(url, 0, 0);
        }

        public Sheet TryToCreateSheet(string url, int horisontalSeparator, int verticalSeparator)
        {
            var sheet = new Sheet()
            {
                HorizontalSeparator = horisontalSeparator,
                VerticalSeparator = verticalSeparator
            };

            try
            {
                sheet.Data = GetData(
                    HttpManager.GetSpreadsheetIdFromUrl(url),
                    HttpManager.GetGidFromUrl(url)
                );
            }
            catch (Exception)
            {
                sheet.Data = new ValueRange();
            }

            return sheet;
        }

        public ValueRange GetData(string spreadsheetId, string gid)
        {
            return sheetsService
                .Spreadsheets
                .Values
                .Get(spreadsheetId, GetSheetName(spreadsheetId, gid))
                .Execute();
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

        public static bool TryToCreateConnect(Stream stream)
        {
            throw new NotImplementedException();
        }
    }
}
