using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Providers;

[PublicAPI]
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
        SpreadsheetsResource.ValuesResource.ClearRequest request =
            _service.Spreadsheets.Values.Clear(new ClearValuesRequest(), SpreadsheetId, range);
        return request.ExecuteAsync();
    }

    internal async Task<IList<IList<object>>> GetValueListAsync(string range, bool formula)
    {
        SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum renderOption = formula
            ? SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMULA
            : SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
        ValueRange valueRange = await GetValuesAsync(range, renderOption);
        return valueRange.Values;
    }

    internal Task UpdateValuesAsync(string range, IList<IList<object>> values)
    {
        ValueRange body = new() { Values = values };
        SpreadsheetsResource.ValuesResource.UpdateRequest request =
            _service.Spreadsheets.Values.Update(body, SpreadsheetId, range);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        return request.ExecuteAsync();
    }

    internal async Task LoadSpreadsheetAsync()
    {
        SpreadsheetsResource.GetRequest request = new(_service, SpreadsheetId);
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
        CopySheetToAnotherSpreadsheetRequest body = new() { DestinationSpreadsheetId = destinationSpreadsheetId };
        SpreadsheetsResource.SheetsResource.CopyToRequest request = new(_service, body, SpreadsheetId, sheetId);
        return request.ExecuteAsync();
    }

    private Task BatchUpdateAsync(IList<Request> requests)
    {
        BatchUpdateSpreadsheetRequest body = new() { Requests = requests };
        SpreadsheetsResource.BatchUpdateRequest request = new(_service, body, SpreadsheetId);
        return request.ExecuteAsync();
    }

    private static Request CreateDeleteSheetRequest(int? sheetId)
    {
        DeleteSheetRequest request = new() { SheetId = sheetId };
        return new Request { DeleteSheet = request };
    }

    private static Request CreateRenameSheetRequest(int? sheetId, string title)
    {
        SheetProperties properties = new()
        {
            SheetId = sheetId,
            Title = title
        };

        UpdateSheetPropertiesRequest request = new()
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
        Spreadsheet body = new() { Properties = properties };
        SpreadsheetsResource.CreateRequest request = new(_service, body);
        return request.ExecuteAsync();
    }

    internal readonly BaseClientService.Initializer ServiceInitializer;
    internal readonly string SpreadsheetId;

    private static readonly string[] Scopes = { SheetsService.Scope.Drive };

    private readonly SheetsService _service;
    private Spreadsheet _spreadsheet;
    private List<Request> _requestsToExecute;
}
