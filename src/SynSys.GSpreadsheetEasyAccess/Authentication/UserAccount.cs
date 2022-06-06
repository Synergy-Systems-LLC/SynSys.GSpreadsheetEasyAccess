using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using SynSys.GSpreadsheetEasyAccess.Authentication.Exceptions;
using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace SynSys.GSpreadsheetEasyAccess.Authentication
{
    /// <summary>
    /// User accounts are managed as Google accounts,
    /// they represent a developer, administrator, or any other person who interacts with Google Cloud.<br/>
    /// These are for scenarios where your application needs to access resources as a human user.<br/>
    /// For more information, see
    /// <a href="https://cloud.google.com/docs/authentication/end-user">Authenticate as an end user</a>.
    /// </summary>
    public class UserAccount : Principal
    {
        private byte[] _credentials;
        private OAuthSheetsScope _scope;

        /// <summary>
        /// Authenticate as a human user.
        /// </summary>
        /// <returns></returns>
        public UserAccount(byte[] credentials, OAuthSheetsScope scope)
        {
            _credentials = credentials;
            _scope = scope;
        }

        /// <summary>
        /// Time to try to connect to the application on the Google Cloud Platform in seconds.<br/>
        /// The default value is <c>30</c>.
        /// </summary>
        public byte CancellationSeconds { get; set; } = 30;

        /// <summary>
        /// Return an object representing Google Sheets service.<br/>
        /// </summary>
        /// <remarks>
        /// It
        /// <a href="https://developers.google.com/resources/api-libraries/documentation/sheets/v4/csharp/latest/classGoogle_1_1Apis_1_1Sheets_1_1v4_1_1SheetsService.html">
        /// Google.Apis.Sheets.v4 library object
        /// </a>,
        /// through which work with Google Sheets API takes place.
        /// </remarks>
        /// <exception cref="AuthenticationTimedOutException"></exception>
        /// <exception cref="UserCanceledAuthenticationException"></exception>
        public override SheetsService GetSheetsService()
        {
            try
            {
                return new SheetsService(
                    new BaseClientService.Initializer
                    {
                        HttpClientInitializer = GetUserCredential(new MemoryStream(_credentials))
                    }
                );
            }
            catch(AggregateException ae)
            {
                var message = "Cancel Authentication";

                foreach (Exception e in ae.InnerExceptions)
                {
                    if (e is TaskCanceledException)
                    {
                        throw new AuthenticationTimedOutException(message, e); 
                    }

                    var ex = e as TokenResponseException;

                    if (ex != null && ex.Error.Error == "access_denied")
                    {
                        throw new UserCanceledAuthenticationException(message, ex);
                    }
                }

                throw;
            }
        }


        private UserCredential GetUserCredential(Stream credentials)
        {
            return GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(credentials).Secrets,
                _scope.Value,
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

        private string GenerateTokenPath()
        {
            var uriPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().EscapedCodeBase);
            var outputDirectory = new Uri(uriPath).LocalPath;
            return Path.Combine(outputDirectory, "token.json");
        }
    }
}
