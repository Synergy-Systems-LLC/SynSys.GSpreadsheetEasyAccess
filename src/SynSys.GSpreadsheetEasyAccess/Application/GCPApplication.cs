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

        #region Authentication

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

        #endregion

        #region Create Sheet

        /// <summary>
        /// Creating Google spreadsheet sheet and get it's representation as an instance of the NativeSheet type.
        /// </summary>
        /// <remarks>
        /// After creating a sheet, you can immediately work with it.
        /// </remarks>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetTitle"></param>
        /// <returns>
        /// NativeSheet is a list of Rows.<br/>
        /// The header is the same as in the Google table in A1 notation.<br/>
        /// Each row has the same number of cells.<br/>
        /// Each cell has a string value and column with number and title.
        /// </returns>
        /// <exception cref="InvalidOperationException"/>
        /// <exception cref="UserAccessDeniedException"/>
        /// <exception cref="SpreadsheetNotFoundException"/>
        /// <exception cref="SheetExistsException"/>
        public NativeSheet CreateNativeSheet(string spreadsheetId, string sheetTitle)
        {
            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            CreateGoogleSheet(spreadsheet, spreadsheetId, sheetTitle);

            var sheetModel = new NativeSheet()
            {
                SpreadsheetTitle = spreadsheet.Properties.Title,
                SpreadsheetId = spreadsheetId,
                Title = sheetTitle,
            };

            return sheetModel;
        }

        /// <summary>
        /// Creating Google spreadsheet sheet and get it's representation as an instance of the UserSheet type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetTitle"></param>
        /// <param name="head"></param>
        /// <remarks>
        /// After creating a sheet, you can immediately work with it.
        /// </remarks>
        /// <returns>
        /// UserSheet is a list of Rows without first row.<br/>
        /// First row is a header of sheet.<br/>
        /// Each row has the same number of cells.<br/>
        /// Each cell has a string value and column with number and title.
        /// </returns>
        /// <exception cref="InvalidOperationException"/>
        /// <exception cref="UserAccessDeniedException"/>
        /// <exception cref="SpreadsheetNotFoundException"/>
        /// <exception cref="InvalidSheetHeadException"/>
        /// <exception cref="SheetExistsException"/>
        public UserSheet CreateUserSheet(string spreadsheetId, string sheetTitle, IEnumerable<string> head)
        {
            if (head.Any() == false)
            {
                throw new InvalidSheetHeadException("Cannot create a sheet with an empty header");
            }

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            CreateGoogleSheet(spreadsheet, spreadsheetId, sheetTitle);
            AddHeadForGoogleSheet(spreadsheetId, sheetTitle, head);

            var sheetModel = new UserSheet()
            {
                SpreadsheetTitle = spreadsheet.Properties.Title,
                SpreadsheetId = spreadsheetId,
                Title = sheetTitle,
            };
            sheetModel.CreateHead(head.ToList());

            return sheetModel;
        }

        /// <summary>
        /// Creating Google spreadsheet sheet and get it's representation as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetTitle"></param>
        /// <param name="head"></param>
        /// <param name="keyName"></param>
        /// <remarks>
        /// After creating a sheet, you can immediately work with it.
        /// </remarks>
        /// <returns>
        /// SheetModel is a list of Rows without first row.<br/>
        /// First row is a header of sheet.<br/>
        /// Each row has the same number of cells and has key column.<br/>
        /// Each cell has a string value and column with number and title.
        /// </returns>
        /// <exception cref="InvalidOperationException"/>
        /// <exception cref="UserAccessDeniedException"/>
        /// <exception cref="SpreadsheetNotFoundException"/>
        /// <exception cref="InvalidSheetHeadException"/>
        /// <exception cref="SheetKeyNotFoundException"/>
        /// <exception cref="SheetExistsException"/>
        public UserSheet CreateUserSheet(string spreadsheetId, string sheetTitle, IEnumerable<string> head, string keyName)
        {
            if (head.Contains(keyName) == false)
            {
                throw new SheetKeyNotFoundException("Cannot create a sheet with a key that is not in the headers");
            }

            var sheetModel = CreateUserSheet(spreadsheetId, sheetTitle, head);
            sheetModel.KeyName = keyName;

            return sheetModel;
        }

        #endregion

        #region Get Sheet

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
        public NativeSheet GetNativeSheet(GoogleSheetUri uri)
        {
            return GetNativeSheet(uri.SpreadsheetId, uri.SheetId);
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
        public NativeSheet GetNativeSheet(string spreadsheetId, int gid)
        {
            CheckSheetService();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            var sheetModel = new NativeSheet()
            {
                SpreadsheetId = spreadsheetId,
                Gid = gid,
                SpreadsheetTitle = spreadsheet.Properties.Title,
                Title = sheet.Properties.Title,
            };

            sheetModel.Fill(data);

            return sheetModel;
        }

        /// <summary>
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetTitle"></param>
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
        public NativeSheet GetNativeSheet(string spreadsheetId, string sheetTitle)
        {
            CheckSheetService();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, sheetTitle);
            IList<IList<object>> data = GetData(spreadsheetId, sheetTitle);

            var sheetModel = new NativeSheet()
            {
                SpreadsheetId = spreadsheetId,
                Title = sheetTitle,
                SpreadsheetTitle = spreadsheet.Properties.Title,
                Gid = sheet.Properties.SheetId.Value,
            };

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
        public UserSheet GetUserSheet(GoogleSheetUri uri)
        {
            return GetUserSheet(uri.SpreadsheetId, uri.SheetId);
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
        public UserSheet GetUserSheet(GoogleSheetUri uri, string keyName)
        {
            return GetUserSheet(uri.SpreadsheetId, uri.SheetId, keyName);
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
        public UserSheet GetUserSheet(string spreadsheetId, int gid)
        {
            CheckSheetService();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            var sheetModel = new UserSheet()
            {
                SpreadsheetId = spreadsheetId,
                Gid = gid,
                SpreadsheetTitle = spreadsheet.Properties.Title,
                Title = sheet.Properties.Title,
                KeyName = string.Empty,
            };

            sheetModel.ValidateData(data);
            sheetModel.Fill(data);

            return sheetModel;
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
        public UserSheet GetUserSheet(string spreadsheetId, int gid, string keyName)
        {
            CheckSheetService();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, gid);
            IList<IList<object>> data = GetData(spreadsheetId, sheet.Properties.Title);

            var sheetModel = new UserSheet()
            {
                SpreadsheetId = spreadsheetId,
                Gid = gid,
                SpreadsheetTitle = spreadsheet.Properties.Title,
                Title = sheet.Properties.Title,
                KeyName = keyName,
            };

            sheetModel.ValidateData(data, keyName);
            sheetModel.Fill(data);

            return sheetModel;
        }

        /// <summary>
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetTitle"></param>
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
        public UserSheet GetUserSheet(string spreadsheetId, string sheetTitle)
        {
            CheckSheetService();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, sheetTitle);
            IList<IList<object>> data = GetData(spreadsheetId, sheetTitle);

            var sheetModel = new UserSheet()
            {
                SpreadsheetId = spreadsheetId,
                Title = sheetTitle,
                SpreadsheetTitle = spreadsheet.Properties.Title,
                Gid = sheet.Properties.SheetId.Value,
                KeyName = string.Empty,
            };

            sheetModel.ValidateData(data);
            sheetModel.Fill(data);

            return sheetModel;
        }

        /// <summary>
        /// Receiving data from a Google spreadsheet sheet as an instance of the SheetModel type.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetTitle"></param>
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
        public UserSheet GetUserSheet(string spreadsheetId, string sheetTitle, string keyName)
        {
            CheckSheetService();

            Spreadsheet spreadsheet = GetGoogleSpreadsheet(spreadsheetId);
            Sheet sheet = GetGoogleSheet(spreadsheet, sheetTitle);
            IList<IList<object>> data = GetData(spreadsheetId, sheetTitle);

            var sheetModel = new UserSheet()
            {
                SpreadsheetId = spreadsheetId,
                Title = sheetTitle,
                SpreadsheetTitle = spreadsheet.Properties.Title,
                Gid = sheet.Properties.SheetId.Value,
                KeyName = keyName,
            };

            sheetModel.ValidateData(data, keyName);
            sheetModel.Fill(data);

            return sheetModel;
        }

        #endregion

        #region Update Sheet

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
        public void UpdateSheet(AbstractSheet sheetModel)
        {
            CheckSheetService();
            CheckPrincipal("Update sheet");

            try
            {
                CreateAppendRowsRequest(sheetModel)?.Execute();

                if (sheetModel is NativeSheet)
                {
                    // TODO пересмотреть логику. Вдруг надо будет вставлять столбец со значениями.
                    CreateAppendEmptyColumnsRequest(sheetModel)?.Execute();
                    CreateInsertEmptyColumnsRequest(sheetModel)?.Execute();
                }
                else
                {
                    // TODO пересмотреть логику
                    CreateAppendColumnWithValuesRequest(sheetModel)?.Execute();
                    CreateInsertEmptyColumnsRequest(sheetModel)?.Execute();
                }

                CreateUpdateCellDataRequest(sheetModel)?.Execute();

                CreateDeleteRowsRequest(sheetModel)?.Execute();
                CreateDeleteColumnsRequest(sheetModel)?.Execute();

                sheetModel.ClearDeletedRows();
                sheetModel.ResetRowStatuses();
                sheetModel.RenumberRows();

                sheetModel.ResetColumnStatuses();
                //sheetModel.TryToClearDeletedColumns();
                //sheetModel.TryToResetColumnStatuses();
            }
            catch (GoogleApiException e) when (e.HttpStatusCode == HttpStatusCode.Forbidden && e.Error.Message.Contains("insufficient authentication scopes"))
            {
                throw new OAuthSheetsScopeException(e.Error.Message, e);
            }
            catch (GoogleApiException e) when (e.HttpStatusCode == HttpStatusCode.Forbidden && e.Error.Message.Contains("does not have permission"))
            {
                throw new UserAccessDeniedException(e.Error.Message, e)
                {
                    Operation = $"Update sheet: {sheetModel.SpreadsheetTitle}/{sheetModel.Title}",
                };
            }
        }

        #endregion

        #region Check Sheet

        /// <summary>
        /// Check the presence of a sheet in the Google spreadsheet by name.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetTitle"></param>
        /// <exception cref="InvalidOperationException"/>
        /// <exception cref="InvalidApiKeyException"/>
        /// <exception cref="UserAccessDeniedException"/>
        /// <exception cref="SpreadsheetNotFoundException"/>
        public bool IsSheetExists(string spreadsheetId, string sheetTitle)
        {
            try
            {
                return IsSheetExists(GetGoogleSpreadsheet(spreadsheetId), sheetTitle);
            }
            catch (SpreadsheetNotFoundException)
            {
                return false;
            }
        }

        /// <summary>
        /// Check the absence of a sheet in the Google spreadsheet by name.
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetTitle"></param>
        /// <exception cref="InvalidOperationException"/>
        /// <exception cref="InvalidApiKeyException"/>
        /// <exception cref="UserAccessDeniedException"/>
        public bool IsSheetNotExists(string spreadsheetId, string sheetTitle)
        {
            return !IsSheetExists(spreadsheetId, sheetTitle);
        }

        #endregion


        #region Check Sheet

        private bool IsSheetExists(Spreadsheet spreadsheet, string sheetTitle)
        {
            return spreadsheet
                .Sheets
                .Select(s => s.Properties.Title)
                .Contains(sheetTitle);
        }

        #endregion

        #region CheckFields

        /// <summary>
        /// 
        /// </summary>
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

        // TODO переназвать метод!
        /// <summary>
        /// 
        /// </summary>
        /// <exception cref="UserAccessDeniedException"></exception>
        private void CheckPrincipal(string operation)
        {
            if (_principal is ServiceAccount)
            {
                throw new UserAccessDeniedException(
                    $"Can't modify Google spreadsheets using {nameof(ServiceAccount)}.")
                {
                    Operation = operation
                };
            }
        }

        #endregion

        #region Google Sheets

        /// <summary>
        /// 
        /// </summary>
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
                    Operation = $"Receiving spreadsheet with id: {spreadsheetId}"
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

        /// <summary>
        /// 
        /// </summary>
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
                    SpreadsheetTitle = spreadsheet.Properties.Title,
                    SpreadsheetId = spreadsheet.SpreadsheetId,
                    SheetGid = gid.ToString()
                };
            }

            return sheet;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <exception cref="SheetNotFoundException"></exception>
        private Sheet GetGoogleSheet(Spreadsheet spreadsheet, string sheetTitle)
        {
            Sheet sheet = spreadsheet.Sheets.ToList().Find(s => s.Properties.Title == sheetTitle);

            if (sheet == null)
            {
                throw new SheetNotFoundException(
                    $"Spreadsheet {spreadsheet.Properties.Title} " +
                    $"with id: {spreadsheet.SpreadsheetId} " +
                    $"not found sheet with name: {sheetTitle}"
                )
                {
                    SpreadsheetTitle = spreadsheet.Properties.Title,
                    SpreadsheetId = spreadsheet.SpreadsheetId,
                    SheetTitle = sheetTitle
                };
            }

            return sheet;
        }

        private IList<IList<object>> GetData(string spreadsheetId, string sheetTitle)
        {
            return _sheetsService
                .Spreadsheets
                .Values
                .Get(spreadsheetId, sheetTitle)
                .Execute()
                .Values ?? new List<IList<object>>();
        }

        #endregion

        #region Requests

        private SpreadsheetsResource.ValuesResource.BatchUpdateRequest CreateUpdateCellDataRequest(AbstractSheet sheet)
        {
            if (sheet.Rows.FindAll(row => row.Status == ChangeStatus.ToChange).Count <= 0)
            {
                return null;
            }
 
            var requestBody = new BatchUpdateValuesRequest
            {
                Data = sheet.GetChangeValueRange(),
                ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW.ToString()
            };

            return _sheetsService
                .Spreadsheets
                .Values
                .BatchUpdate(requestBody, sheet.SpreadsheetId);
        }

        private SpreadsheetsResource.ValuesResource.AppendRequest CreateAppendRowsRequest(AbstractSheet sheet)
        {
            // TODO привести два метода AppendRequest в соответствие
            if (sheet.Rows.FindAll(row => row.Status == ChangeStatus.ToAppend).Count <= 0)
            {
                return null;
            }
 
            var request = _sheetsService
                .Spreadsheets
                .Values
                .Append(
                    sheet.GetAppendValueRange(),
                    sheet.SpreadsheetId,
                    sheet.Title
                );

            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            return request;
        }

        private SpreadsheetsResource.ValuesResource.AppendRequest CreateAppendColumnWithValuesRequest(AbstractSheet sheet)
        {
            AbstractColumn column = sheet.GetFirstAddedColumn();

            if (column == null)
            {
                return null;
            }

            List<IList<object>> values = sheet.GetAddedColumnsValues();

            var request = _sheetsService
                .Spreadsheets
                .Values
                .Append(
                    new ValueRange()
                    {
                        Values = values,
                    },
                    sheet.SpreadsheetId,
                    $"{sheet.Title}!{NativeColumn.GenerateA1NotationTitle(column.Number)}1"
                );

            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            return request;
        }

        private SpreadsheetsResource.BatchUpdateRequest CreateDeleteRowsRequest(AbstractSheet sheet)
        {
            if (sheet.Rows.FindAll(row => row.Status == ChangeStatus.ToDelete).Count <= 0)
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
                        "ROWS",  // TODO сделать Enum-о подобный класс
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

        private SpreadsheetsResource.BatchUpdateRequest CreateAppendEmptyColumnsRequest(AbstractSheet sheet)
        {
            int amountAppendColumns = sheet.GetCountAppendColumns();

            if (amountAppendColumns == 0)
            {
                return null;
            }

            var request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>
                {
                    new Request()
                    {
                        AppendDimension = new AppendDimensionRequest()
                        {
                            SheetId = sheet.Gid,
                            Dimension = "COLUMNS",
                            Length = amountAppendColumns
                        }
                    }
                }
            };

            return _sheetsService.Spreadsheets.BatchUpdate(request, sheet.SpreadsheetId);
        }

        private SpreadsheetsResource.BatchUpdateRequest CreateInsertEmptyColumnsRequest(AbstractSheet sheet)
        {
            List<List<AbstractColumn>> groupOfColumns = sheet.GetInsertColumnsGroups();

            if (groupOfColumns.Count == 0)
            {
                return null;
            }

            var requests = new List<Request>();

            foreach (List<AbstractColumn> group in groupOfColumns)
            {
                requests.Add(
                    new Request()
                    {
                        InsertDimension = new InsertDimensionRequest()
                        {
                            Range = new DimensionRange()
                            {
                                SheetId = sheet.Gid,
                                Dimension = "COLUMNS",
                                StartIndex = group.First().Number - 1,
                                EndIndex = group.Last().Number
                            },
                            InheritFromBefore = false
                        }
                    }
                );
            }

            var request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = requests
            };

            return _sheetsService.Spreadsheets.BatchUpdate(request, sheet.SpreadsheetId);
        }

        private SpreadsheetsResource.BatchUpdateRequest CreateDeleteColumnsRequest(AbstractSheet sheet)
        {
            var requestBody = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>()
            };

            foreach (List<int> groupNumbers in sheet.GetDeleteColumns())
            {
                if (!groupNumbers.Any())
                {
                    continue;
                }

                requestBody.Requests.Add(
                    CreateDeleteDimensionRequest(
                        sheet.Gid,
                        "COLUMNS",
                        groupNumbers.Last() - 1,
                        groupNumbers.First()
                    )
                );
            }

            // TODO вот это скорей всего не будет,
            // надо делать проверку в самом начале этого метода по данным из SheetModel
            if (!requestBody.Requests.Any())
            {
                return null;
            }

            var deleteRequest =  new SpreadsheetsResource.BatchUpdateRequest(
                _sheetsService,
                requestBody,
                sheet.SpreadsheetId
            );

            return deleteRequest;
        }

        private SpreadsheetsResource.BatchUpdateRequest CreateMoveColumnsRequest(AbstractSheet sheet)
        {
            // TODO дописать получение перемещаемых столбцов из листа

            var request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>
                {
                    new Request
                    {
                        MoveDimension = new MoveDimensionRequest()
                        {
                            Source = new DimensionRange()
                            {
                                Dimension = "COLUMNS",
                                StartIndex = 1, // TODO получить из sheet
                                EndIndex = 1 // TODO получить из sheet
                            },
                            DestinationIndex = 1 // TODO получить из sheet 
                        }
                    }
                }
            };

            return _sheetsService.Spreadsheets.BatchUpdate(request, sheet.SpreadsheetId);
        }

        private SpreadsheetsResource.BatchUpdateRequest CreateAddSheetRequest(string spreadsheetId, string sheetTitle)
        {
            var batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        AddSheet = new AddSheetRequest()
                        {
                            Properties = new SheetProperties()
                            {
                                Title = sheetTitle,
                            }
                        }
                    }
                }
            };

            return _sheetsService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, spreadsheetId);
        }

        private Request CreateDeleteDimensionRequest(int gid, string dimension, int startRow, int endRow)
        {
            return new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        SheetId = gid,
                        Dimension = dimension,
                        StartIndex = startRow,
                        EndIndex = endRow,
                    }
                }
            };
        }

        /// <summary>
        /// 
        /// </summary>
        /// <exception cref="UserAccessDeniedException"></exception>
        /// <exception cref="SheetExistsException"></exception>
        /// <exception cref="CreatingSheetException"></exception>
        private void CreateGoogleSheet(Spreadsheet spreadsheet, string spreadsheetId, string sheetTitle)
        {
            CheckSheetService();
            CheckPrincipal("Create sheet");

            if (IsSheetExists(spreadsheet, sheetTitle))
            {
                throw new SheetExistsException()
                {
                    SpreadsheetId = spreadsheetId,
                    SpreadsheetTitle = spreadsheet.Properties.Title,
                    SheetTitle = sheetTitle
                };
            }

            try
            {
                CreateAddSheetRequest(spreadsheetId, sheetTitle).Execute();
            }
            catch (Exception e)
            {
                throw new CreatingSheetException("Couldn't add sheet to google spreadsheet", e)
                {
                    SpreadsheetId = spreadsheetId,
                    SpreadsheetTitle = spreadsheet.Properties.Title,
                    SheetTitle = sheetTitle,
                };
            }
        }

        private void AddHeadForGoogleSheet(string spreadsheetId, string title, IEnumerable<string> head)
        {
            var range = new ValueRange
            {
                Values = new List<IList<object>>()
                {
                    head.Cast<object>().ToList()
                }
            };

            var request = _sheetsService
                .Spreadsheets
                .Values
                .Append(range, spreadsheetId, title);

            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            request.Execute();
        }

        #endregion
    }
}
