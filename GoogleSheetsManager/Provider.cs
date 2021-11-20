using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsManager
{
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    [SuppressMessage("ReSharper", "ClassNeverInstantiated.Global")]
    public sealed class Provider : IDisposable
    {
        public Provider(string credentialJson, string applicationName, string sheetId)
            : this(GoogleCredential.FromJson(credentialJson).CreateScoped(Scopes), applicationName, sheetId)
        {
        }

        private Provider(IConfigurableHttpClientInitializer httpClientInitializer, string applicationName,
            Spreadsheet spreadsheet)
            : this(httpClientInitializer, applicationName, spreadsheet.SpreadsheetId, spreadsheet)
        {
        }

        private Provider(IConfigurableHttpClientInitializer httpClientInitializer, string applicationName, string sheetId,
            Spreadsheet spreadsheet = null)
        {
            var initializer = new BaseClientService.Initializer
            {
                HttpClientInitializer = httpClientInitializer,
                ApplicationName = applicationName
            };

            _sheetsService = new SheetsService(initializer);
            _driveService = new DriveService(initializer);
            _sheetId = sheetId;
            _spreadsheet = spreadsheet;
        }

        public void Dispose()
        {
            _sheetsService.Dispose();
            _driveService.Dispose();
        }

        public void ClearValues(string range)
        {
            var body = new ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest request =
                _sheetsService.Spreadsheets.Values.Clear(body, _sheetId, range);
            request.Execute();
        }

        public string CopyFor(string ownerEmail)
        {
            _spreadsheet = GetSheet();

            Spreadsheet newSpreadSheet = CreateNew(_spreadsheet.Properties);

            using (var provider =
                new Provider(_driveService.HttpClientInitializer, _driveService.ApplicationName, newSpreadSheet))
            {
                CopyContent(this, provider);

                provider.SetupPermissionsFor(ownerEmail);
            }

            return newSpreadSheet.SpreadsheetId;
        }

        internal IEnumerable<IList<object>> GetValues(string range,
            SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum valueRenderOption)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetsService.Spreadsheets.Values.Get(_sheetId, range);
            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption =
                SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
            ValueRange response = request.Execute();
            return response.Values;
        }

        internal void UpdateValues(string range, IList<IList<object>> values)
        {
            var valueRange = new ValueRange { Values = values };
            SpreadsheetsResource.ValuesResource.UpdateRequest request =
                _sheetsService.Spreadsheets.Values.Update(valueRange, _sheetId, range);
            request.ValueInputOption =
                SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            request.Execute();
        }

        private static void CopyContent(Provider from, Provider to)
        {
            List<Sheet> oldSheets = to._spreadsheet.Sheets.ToList();

            Dictionary<string, SheetProperties> newSheetsProperties = CopySheets(from, to);

            IEnumerable<Request> deleteRequests = oldSheets.Select(s => CreateDeleteSheetRequest(s.Properties.SheetId));

            IEnumerable<Request> renameRequests =
                newSheetsProperties.Select(p => CreateRenameSheetRequest(p.Value.SheetId, p.Key));

            List<Request> requests = deleteRequests.Concat(renameRequests).ToList();

            to.BatchUpdate(requests);
        }

        private static Dictionary<string, SheetProperties> CopySheets(Provider from, Provider to)
        {
            var result = new Dictionary<string, SheetProperties>();
            for (int i = 0; i < from._spreadsheet.Sheets.Count; ++i)
            {
                SheetProperties properties = from.CopyTo(to._sheetId, i);
                Sheet sheet = from._spreadsheet.Sheets[i];
                result[sheet.Properties.Title] = properties;
            }
            return result;
        }

        private void SetupPermissionsFor(string ownerEmail)
        {
            IList<Permission> oldPermissions = GetPermission().Permissions;

            AddPermissionTo("owner", "user", ownerEmail);

            AddPermissionTo("writer", "anyone");

            foreach (Permission permission in oldPermissions)
            {
                DowngradePermission(permission.Id);
            }
        }

        private Spreadsheet GetSheet()
        {
            var request = new SpreadsheetsResource.GetRequest(_sheetsService, _sheetId);
            return request.Execute();
        }

        private SheetProperties CopyTo(string destinationSpreadsheetId, int sheetId)
        {
            var body = new CopySheetToAnotherSpreadsheetRequest { DestinationSpreadsheetId = destinationSpreadsheetId };
            var request = new SpreadsheetsResource.SheetsResource.CopyToRequest(_sheetsService, body, _sheetId, sheetId);
            return request.Execute();
        }

        private void BatchUpdate(IList<Request> requests)
        {
            var body = new BatchUpdateSpreadsheetRequest { Requests = requests };
            var request = new SpreadsheetsResource.BatchUpdateRequest(_sheetsService, body, _sheetId);
            request.Execute();
        }

        private Spreadsheet CreateNew(SpreadsheetProperties properties)
        {
            var body = new Spreadsheet { Properties = properties };
            var request = new SpreadsheetsResource.CreateRequest(_sheetsService, body);
            return request.Execute();
        }

        private PermissionList GetPermission()
        {
            PermissionsResource.ListRequest request = _driveService.Permissions.List(_sheetId);
            return request.Execute();
        }

        private void AddPermissionTo(string role, string type, string emailAddress = null)
        {
            var body = new Permission
            {
                Type = type,
                Role = role,
                EmailAddress = emailAddress
            };

            bool owner = role == "owner";

            PermissionsResource.CreateRequest request = _driveService.Permissions.Create(body, _sheetId);
            request.TransferOwnership = owner;
            request.MoveToNewOwnersRoot = owner;

            request.Execute();
        }

        private void DowngradePermission(string id)
        {
            var body = new Permission { Role = "reader" };
            PermissionsResource.UpdateRequest request = _driveService.Permissions.Update(body, _sheetId, id);
            request.Execute();
        }

        private static Request CreateDeleteSheetRequest(int? sheetId)
        {
            var request = new DeleteSheetRequest { SheetId = sheetId };
            return new Request { DeleteSheet = request };
        }

        private static Request CreateRenameSheetRequest(int? sheetId, string title)
        {
            var properties = new SheetProperties
            {
                SheetId = sheetId,
                Title = title
            };

            var request = new UpdateSheetPropertiesRequest
            {
                Fields = "title",
                Properties = properties
            };

            return new Request { UpdateSheetProperties = request };
        }

        private static readonly string[] Scopes = { SheetsService.Scope.Drive };

        private readonly SheetsService _sheetsService;
        private readonly DriveService _driveService;
        private readonly string _sheetId;

        private Spreadsheet _spreadsheet;
    }
}
