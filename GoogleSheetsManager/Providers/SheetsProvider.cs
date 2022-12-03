using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GryphonUtilities;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Providers;

[PublicAPI]
public sealed class SheetsProvider : IDisposable
{
    public SheetsProvider(IConfigGoogleSheets config, string spreadsheetId)
        : this(config.GetCredentialJson(), config.ApplicationName, new TimeManager(config.TimeZoneId), spreadsheetId)
    {
    }

    public SheetsProvider(string credentialJson, string applicationName, TimeManager timeManager, string spreadsheetId)
        : this(CreateInitializer(credentialJson, applicationName), timeManager, spreadsheetId)
    {
    }

    internal SheetsProvider(BaseClientService.Initializer initializer, SheetsService service, TimeManager timeManager,
        string spreadsheetId)
    {
        ServiceInitializer = initializer;
        Service = service;
        TimeManager = timeManager;
        SpreadsheetId = spreadsheetId;
    }

    private SheetsProvider(BaseClientService.Initializer initializer, TimeManager timeManager, string spreadsheetId)
        : this(initializer, new SheetsService(initializer), timeManager, spreadsheetId)
    {
    }

    public void Dispose() => Service.Dispose();

    public Task ClearValuesAsync(string range)
    {
        SpreadsheetsResource.ValuesResource.ClearRequest request =
            Service.Spreadsheets.Values.Clear(new ClearValuesRequest(), SpreadsheetId, range);
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
            Service.Spreadsheets.Values.Update(body, SpreadsheetId, range);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        return request.ExecuteAsync();
    }

    internal Task<Spreadsheet> LoadSpreadsheetAsync()
    {
        SpreadsheetsResource.GetRequest request = Service.Spreadsheets.Get(SpreadsheetId);
        request.IncludeGridData = true;
        return request.ExecuteAsync();
    }

    internal async Task CopyContentAndPlanToRenameSheetsAsync(SheetsProvider from, Spreadsheet spreadsheet)
    {
        foreach (Sheet sheet in spreadsheet.Sheets.Where(s => s.Properties.SheetId.HasValue))
        {
            // ReSharper disable once NullableWarningSuppressionIsUsed
            //   sheet.Properties.SheetId is null-checked already
            SheetProperties properties = await from.CopyToAsync(SpreadsheetId, sheet.Properties.SheetId!.Value);
            Request renameRequest = CreateRenameSheetRequest(properties.SheetId, sheet.Properties.Title);
            _requestsToExecute.Add(renameRequest);
        }
    }

    internal Task ExecutePlannedAsync()
    {
        BatchUpdateSpreadsheetRequest body = new() { Requests = _requestsToExecute };
        SpreadsheetsResource.BatchUpdateRequest request = new(Service, body, SpreadsheetId);
        return request.ExecuteAsync();
    }

    private Task<SheetProperties> CopyToAsync(string destinationSpreadsheetId, int sheetId)
    {
        CopySheetToAnotherSpreadsheetRequest body = new() { DestinationSpreadsheetId = destinationSpreadsheetId };
        SpreadsheetsResource.SheetsResource.CopyToRequest request = new(Service, body, SpreadsheetId, sheetId);
        return request.ExecuteAsync();
    }

    internal Task RenameSheetAsync(int? sheetId, string title)
    {
        Request request = CreateRenameSheetRequest(sheetId, title);
        BatchUpdateSpreadsheetRequest body = new() { Requests = request.WrapWithList() };
        SpreadsheetsResource.BatchUpdateRequest batchRequest = Service.Spreadsheets.BatchUpdate(body, SpreadsheetId);
        return batchRequest.ExecuteAsync();
    }

    internal Task<Spreadsheet> CreateNewSpreadsheetAsync(SpreadsheetProperties properties)
    {
        Spreadsheet body = new() { Properties = properties };
        SpreadsheetsResource.CreateRequest request = new(Service, body);
        return request.ExecuteAsync();
    }

    internal void PlanToDeleteSheets(Spreadsheet spreadsheet)
    {
        _requestsToExecute.Clear();
        _requestsToExecute.AddRange(spreadsheet.Sheets.Select(s => CreateDeleteSheetRequest(s.Properties.SheetId)));
    }

    private static Request CreateDeleteSheetRequest(int? sheetId)
    {
        DeleteSheetRequest request = new() { SheetId = sheetId };
        return new Request { DeleteSheet = request };
    }

    private readonly List<Request> _requestsToExecute = new();

    private static BaseClientService.Initializer CreateInitializer(string credentialJson, string applicationName)
    {
        GoogleCredential credential = GoogleCredential.FromJson(credentialJson).CreateScoped(Scopes);
        return new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = applicationName
        };
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

    private Task<ValueRange> GetValuesAsync(string range,
        SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum valueRenderOption)
    {
        SpreadsheetsResource.ValuesResource.GetRequest request =
            Service.Spreadsheets.Values.Get(SpreadsheetId, range);
        request.ValueRenderOption = valueRenderOption;
        request.DateTimeRenderOption =
            SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
        return request.ExecuteAsync();
    }

    internal readonly BaseClientService.Initializer ServiceInitializer;
    internal readonly string SpreadsheetId;
    internal readonly SheetsService Service;
    internal readonly TimeManager TimeManager;

    private static readonly string[] Scopes = { SheetsService.Scope.Drive };
}
