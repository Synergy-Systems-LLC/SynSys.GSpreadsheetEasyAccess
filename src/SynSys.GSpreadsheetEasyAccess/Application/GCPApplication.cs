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
    /// Represents an application on the Google Cloud Platform that has access
    /// to Google Sheets API.
    /// <a href="https://developers.google.com/sheets/api?hl=en_US">Google Sheets API</a>.
    /// </summary>
    /// <remarks>
    /// The class serves to receive and update data from Google Sheets.<br/>
    /// Methods can only be used after successful authentication.
    /// </remarks>
    public class GCPApplication
    {
        private SheetsService _sheetsService;
        private Principal _principal;

        /// <summary>
        /// To gain access to the Google Sheets API, you must be authenticated.
        /// It is necessary to specify who is authenticating.
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
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="uri">Full uri of sheet</param>
        /// <returns>
        /// SheetModel is a list of Rows.<br/>
        /// Header is absent.<br/>
        /// Each row has the same number of cells.<br/>
        /// Each cell has a string value.
        /// </returns>
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
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="gid"></param>
        /// <returns>
        /// SheetModel is a list of Rows.<br/>
        /// Header is absent.<br/>
        /// Each row has the same number of cells.<br/>
        /// Each cell has a string value.
        /// </returns>
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
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetName"></param>
        /// <returns>
        /// SheetModel is a list of Rows.<br/>
        /// Header is absent.<br/>
        /// Each row has the same number of cells.<br/>
        /// Each cell has a string value.
        /// </returns>
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
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <returns>
        /// SheetModel is a list of Rows without first row.<br/>
        /// First row is a header of sheet.<br/>
        /// Each row has the same number of cells.<br/>
        /// Each cell has a string value and title, 
        /// which matches the column heading for the given cell.
        /// </returns>
        /// <param name="uri">Full sheet uri</param>
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
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="gid"></param>
        /// <returns>
        /// SheetModel is a list of Rows without first row.<br/>
        /// First row is a header of sheet.<br/>
        /// Each row has the same number of cells.<br/>
        /// Each cell has a string value and title, 
        /// which matches the column heading for the given cell.
        /// </returns>
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
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetName"></param>
        /// <returns>
        /// SheetModel is a list of Rows without first row.<br/>
        /// First row is a header of sheet.<br/>
        /// Each row has the same number of cells.<br/>
        /// Each cell has a string value and title, 
        /// which matches the column heading for the given cell.
        /// </returns>
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
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="uri"></param>
        /// <param name="keyName"></param>
        /// <returns>
        /// SheetModel is a list of Rows without first row.<br/>
        /// First row is a header of sheet.<br/>
        /// Each row has the same number of cells and has key column.<br/>
        /// Each cell has a string value and title, 
        /// which matches the column heading for the given cell.
        /// </returns>
        /// <exception cref="InvalidOperationException"></exception>
        /// <exception cref="InvalidApiKeyException"></exception>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SpreadsheetNotFoundException"></exception>
        /// <exception cref="SheetNotFoundException"></exception>
        /// <exception cref="SheetKeyNotFoundException"></exception>
        /// <exception cref="EmptySheetException"></exception>
        public SheetModel GetSheetWithHeadAndKey(string uri, string keyName)
        {
            return GetSheetWithHeadAndKey(
                HttpUtils.GetSpreadsheetIdFromUri(uri),
                HttpUtils.GetGidFromUri(uri),
                keyName
            );
        }

        /// <summary>
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="gid"></param>
        /// <param name="keyName"></param>
        /// <returns>
        /// SheetModel is a list of Rows without first row.<br/>
        /// First row is a header of sheet.<br/>
        /// Each row has the same number of cells and has key column.<br/>
        /// Each cell has a string value and title, 
        /// which matches the column heading for the given cell.
        /// </returns>
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
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetName"></param>
        /// <param name="keyName"></param>
        /// <returns>
        /// SheetModel is a list of Rows without first row.<br/>
        /// First row is a header of sheet.<br/>
        /// Each row has the same number of cells and has key column.<br/>
        /// Each cell has a string value and title, 
        /// which matches the column heading for the given cell.
        /// </returns>
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
        /// Update the Google spreadsheet sheet based on the modified instance of the SheetModel type.
        /// </summary>
        /// <remarks>
        /// The method changes the data in the cells,
        /// adds rows to the end of the sheet and removes the selected rows.<br />
        /// All these actions are based on requests to Google.
        /// </remarks>
        /// <param name="sheetModel">Google spreadsheet sheet model</param>
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
                    $"Need to authenticate before using this method."
                );
            }
        }

        /// <exception cref="UserAccessDeniedException"></exception>
        private void CheckPrincipal()
        {
            if (_principal is ServiceAccount)
            {
                throw new UserAccessDeniedException(
                    $"Can't modify Google spreadsheets using {nameof(ServiceAccount)}.")
                {
                    Operation = "Sheet update"
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
                // Controversial and implicit decision,
                // but so far this state only appeared due to an invalid API key.
                // Maybe look for a clarifying error state.
                throw new InvalidApiKeyException($"Failed to get spreadsheet with id: {spreadsheetId}", e);
            }
            catch (GoogleApiException e) when (e.HttpStatusCode == HttpStatusCode.Forbidden && e.Error.Message.Contains("does not have permission"))
            {
                throw new UserAccessDeniedException(e.Error.Message, e)
                {
                    Operation = $"Recieving spreadsheet with id: {spreadsheetId}"
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
                    $"Spreadsheet {spreadsheet.Properties.Title} " +
                    $"with id: {spreadsheet.SpreadsheetId} " +
                    $"not found sheet with gid: {gid}"
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
                    $"Spreadsheet {spreadsheet.Properties.Title} " +
                    $"with id: {spreadsheet.SpreadsheetId} " +
                    $"not found sheet with name: {sheetName}"
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
