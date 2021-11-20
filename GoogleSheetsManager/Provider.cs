using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;
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

        public Task ClearValuesAsync(string range)
        {
            var body = new ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest request =
                _sheetsService.Spreadsheets.Values.Clear(body, _sheetId, range);
            return request.ExecuteAsync();
        }

        public async Task<string> CopyForAsync(string name, string folderId, string ownerEmail)
        {
            _spreadsheet = await GetSheetAsync();

            Spreadsheet newSpreadSheet = await CreateNewAsync(_spreadsheet.Properties);

            using (var provider =
                new Provider(_driveService.HttpClientInitializer, _driveService.ApplicationName, newSpreadSheet))
            {
                await CopyContentAsync(this, provider);

                await provider.SetupPermissionsForAsync(ownerEmail);

                await provider.MoveAndRenameAsync(name, folderId);
            }

            return newSpreadSheet.SpreadsheetId;
        }

        internal Task<ValueRange> GetValuesAsync(string range,
            SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum valueRenderOption)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = _sheetsService.Spreadsheets.Values.Get(_sheetId, range);
            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption =
                SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
            return request.ExecuteAsync();
        }

        internal Task UpdateValuesAsync(string range, IList<IList<object>> values)
        {
            var valueRange = new ValueRange { Values = values };
            SpreadsheetsResource.ValuesResource.UpdateRequest request =
                _sheetsService.Spreadsheets.Values.Update(valueRange, _sheetId, range);
            request.ValueInputOption =
                SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            return request.ExecuteAsync();
        }

        private static async Task CopyContentAsync(Provider from, Provider to)
        {
            List<Sheet> oldSheets = to._spreadsheet.Sheets.ToList();

            Dictionary<string, SheetProperties> newSheetsProperties = await CopySheetsAsync(from, to);

            IEnumerable<Request> deleteRequests = oldSheets.Select(s => CreateDeleteSheetRequest(s.Properties.SheetId));

            IEnumerable<Request> renameRequests =
                newSheetsProperties.Select(p => CreateRenameSheetRequest(p.Value.SheetId, p.Key));

            List<Request> requests = deleteRequests.Concat(renameRequests).ToList();

            await to.BatchUpdateAsync(requests);
        }

        private static async Task<Dictionary<string, SheetProperties>> CopySheetsAsync(Provider from, Provider to)
        {
            var result = new Dictionary<string, SheetProperties>();
            for (int i = 0; i < from._spreadsheet.Sheets.Count; ++i)
            {
                SheetProperties properties = await from.CopyToAsync(to._sheetId, i);
                Sheet sheet = from._spreadsheet.Sheets[i];
                result[sheet.Properties.Title] = properties;
            }
            return result;
        }

        private async Task SetupPermissionsForAsync(string ownerEmail)
        {
            PermissionList oldPermissionList = await GetPermissionAsync();
            IList<Permission> oldPermissions = oldPermissionList.Permissions;

            await AddPermissionToAsync("owner", "user", ownerEmail);

            await AddPermissionToAsync("writer", "anyone");

            foreach (Permission permission in oldPermissions)
            {
                await DowngradePermissionAsync(permission.Id);
            }
        }

        private async Task MoveAndRenameAsync(string name, string folderId)
        {
            File file = await GetFileWithParentsAsync();

            string oldParents = string.Join(',', file.Parents);

            await MoveAndRenameAsync(name, folderId, oldParents);
        }

        private async Task MoveAndRenameAsync(string name, string folderId, string oldParents)
        {
            var body = new File { Name = name };

            FilesResource.UpdateRequest request = _driveService.Files.Update(body, _sheetId);
            request.AddParents = folderId;
            request.RemoveParents = oldParents;

            await request.ExecuteAsync();
        }

        private Task<Spreadsheet> GetSheetAsync()
        {
            var request = new SpreadsheetsResource.GetRequest(_sheetsService, _sheetId);
            return request.ExecuteAsync();
        }

        private Task<SheetProperties> CopyToAsync(string destinationSpreadsheetId, int sheetId)
        {
            var body = new CopySheetToAnotherSpreadsheetRequest { DestinationSpreadsheetId = destinationSpreadsheetId };
            var request = new SpreadsheetsResource.SheetsResource.CopyToRequest(_sheetsService, body, _sheetId, sheetId);
            return request.ExecuteAsync();
        }

        private Task BatchUpdateAsync(IList<Request> requests)
        {
            var body = new BatchUpdateSpreadsheetRequest { Requests = requests };
            var request = new SpreadsheetsResource.BatchUpdateRequest(_sheetsService, body, _sheetId);
            return request.ExecuteAsync();
        }

        private Task<Spreadsheet> CreateNewAsync(SpreadsheetProperties properties)
        {
            var body = new Spreadsheet { Properties = properties };
            var request = new SpreadsheetsResource.CreateRequest(_sheetsService, body);
            return request.ExecuteAsync();
        }

        private Task<PermissionList> GetPermissionAsync()
        {
            PermissionsResource.ListRequest request = _driveService.Permissions.List(_sheetId);
            return request.ExecuteAsync();
        }

        private Task AddPermissionToAsync(string role, string type, string emailAddress = null)
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

            return request.ExecuteAsync();
        }

        private Task DowngradePermissionAsync(string id)
        {
            var body = new Permission { Role = "reader" };
            PermissionsResource.UpdateRequest request = _driveService.Permissions.Update(body, _sheetId, id);
            return request.ExecuteAsync();
        }

        private Task<File> GetFileWithParentsAsync()
        {
            FilesResource.GetRequest request = _driveService.Files.Get(_sheetId);
            request.Fields = "parents";
            return request.ExecuteAsync();
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
