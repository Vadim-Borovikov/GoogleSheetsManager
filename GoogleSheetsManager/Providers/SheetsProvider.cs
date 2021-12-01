using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsManager.Providers
{
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    [SuppressMessage("ReSharper", "ClassNeverInstantiated.Global")]
    public sealed class SheetsProvider : IDisposable
    {
        public SheetsProvider(string credentialJson, string applicationName, string spreadsheetId)
            : this(CreateInitializer(credentialJson, applicationName), spreadsheetId)
        {
        }

        private SheetsProvider(BaseClientService.Initializer serviceInitializer, Spreadsheet spreadsheet)
            : this(serviceInitializer, spreadsheet.SpreadsheetId)
        {
            _spreadsheet = spreadsheet;
        }

        private SheetsProvider(BaseClientService.Initializer initializer, string spreadsheetId)
        {
            ServiceInitializer = initializer;
            _service = new SheetsService(initializer);
            SpreadsheetId = spreadsheetId;
        }

        public void Dispose() => _service.Dispose();

        public Task ClearValuesAsync(string range)
        {
            var body = new ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest request =
                _service.Spreadsheets.Values.Clear(body, SpreadsheetId, range);
            return request.ExecuteAsync();
        }

        internal async Task<IList<IList<object>>> GetValueListAsync(string range)
        {
            ValueRange valueRange = await GetValuesAsync(range,
                SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE);
            return valueRange.Values;
        }

        internal Task UpdateValuesAsync(string range, IList<IList<object>> values)
        {
            var body = new ValueRange { Values = values };
            SpreadsheetsResource.ValuesResource.UpdateRequest request =
                _service.Spreadsheets.Values.Update(body, SpreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            return request.ExecuteAsync();
        }

        internal async Task LoadSpreadsheetAsync()
        {
            var request = new SpreadsheetsResource.GetRequest(_service, SpreadsheetId);
            _spreadsheet = await request.ExecuteAsync();
        }

        internal async Task<SheetsProvider> CreateNewWithPropertiesAsync()
        {
            Spreadsheet spreadsheet = await CreateNewAsync(_spreadsheet.Properties);
            return new SheetsProvider(ServiceInitializer, spreadsheet);
        }

        internal void PlanToDeleteSheets()
        {
            _requestsToExecute?.Clear();
            _requestsToExecute = _spreadsheet.Sheets.Select(s => CreateDeleteSheetRequest(s.Properties.SheetId)).ToList();
        }

        internal async Task CopyContentAndPlanToRenameSheetsAsync(SheetsProvider from)
        {
            foreach (Sheet sheet in from._spreadsheet.Sheets.Where(s => s.Properties.SheetId.HasValue))
            {
                // ReSharper disable once PossibleInvalidOperationException
                SheetProperties properties = await from.CopyToAsync(SpreadsheetId, sheet.Properties.SheetId.Value);
                Request renameRequest = CreateRenameSheetRequest(properties.SheetId, sheet.Properties.Title);
                _requestsToExecute.Add(renameRequest);
            }
        }

        internal Task ExecutePlanned() => BatchUpdateAsync(_requestsToExecute);

        private Task<SheetProperties> CopyToAsync(string destinationSpreadsheetId, int sheetId)
        {
            var body = new CopySheetToAnotherSpreadsheetRequest { DestinationSpreadsheetId = destinationSpreadsheetId };
            var request = new SpreadsheetsResource.SheetsResource.CopyToRequest(_service, body, SpreadsheetId, sheetId);
            return request.ExecuteAsync();
        }

        private Task BatchUpdateAsync(IList<Request> requests)
        {
            var body = new BatchUpdateSpreadsheetRequest { Requests = requests };
            var request = new SpreadsheetsResource.BatchUpdateRequest(_service, body, SpreadsheetId);
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

        private static BaseClientService.Initializer CreateInitializer(string credentialJson, string applicationName)
        {
            GoogleCredential credential = GoogleCredential.FromJson(credentialJson).CreateScoped(Scopes);
            return new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName
            };
        }

        private Task<ValueRange> GetValuesAsync(string range,
            SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum valueRenderOption)
        {
            SpreadsheetsResource.ValuesResource.GetRequest request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
            request.ValueRenderOption = valueRenderOption;
            request.DateTimeRenderOption =
                SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
            return request.ExecuteAsync();
        }

        private Task<Spreadsheet> CreateNewAsync(SpreadsheetProperties properties)
        {
            var body = new Spreadsheet { Properties = properties };
            var request = new SpreadsheetsResource.CreateRequest(_service, body);
            return request.ExecuteAsync();
        }

        internal readonly BaseClientService.Initializer ServiceInitializer;
        internal readonly string SpreadsheetId;

        private static readonly string[] Scopes = { SheetsService.Scope.Drive };

        private readonly SheetsService _service;
        private Spreadsheet _spreadsheet;
        private List<Request> _requestsToExecute;
    }
}
