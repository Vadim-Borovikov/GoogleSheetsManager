using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Providers;

[PublicAPI]
public class SheetsProvider : IDisposable
{
    public SheetsProvider(string credentialJson, string applicationName, string spreadsheetId)
        : this(CreateInitializer(credentialJson, applicationName), spreadsheetId)
    {
    }

    protected SheetsProvider(BaseClientService.Initializer initializer, SheetsService service, string spreadsheetId)
    {
        ServiceInitializer = initializer;
        Service = service;
        SpreadsheetId = spreadsheetId;
    }

    private SheetsProvider(BaseClientService.Initializer initializer, string spreadsheetId)
        : this(initializer, new SheetsService(initializer), spreadsheetId)
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

    internal Task<Spreadsheet> LoadSpreadsheet()
    {
        SpreadsheetsResource.GetRequest request = Service.Spreadsheets.Get(SpreadsheetId);
        request.IncludeGridData = true;
        return request.ExecuteAsync();
    }

    internal Task RenameSheetAsync(int? sheetId, string title)
    {
        Request request = CreateRenameSheetRequest(sheetId, title);
        List<Request> requests = new() { request };
        BatchUpdateSpreadsheetRequest body = new() { Requests = requests };
        SpreadsheetsResource.BatchUpdateRequest batchRequest = Service.Spreadsheets.BatchUpdate(body, SpreadsheetId);
        return batchRequest.ExecuteAsync();
    }

    protected static BaseClientService.Initializer CreateInitializer(string credentialJson, string applicationName)
    {
        GoogleCredential credential = GoogleCredential.FromJson(credentialJson).CreateScoped(Scopes);
        return new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = applicationName
        };
    }

    protected static Request CreateRenameSheetRequest(int? sheetId, string title)
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
        SpreadsheetsResource.ValuesResource.GetRequest request = Service.Spreadsheets.Values.Get(SpreadsheetId, range);
        request.ValueRenderOption = valueRenderOption;
        request.DateTimeRenderOption =
            SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
        return request.ExecuteAsync();
    }

    internal readonly BaseClientService.Initializer ServiceInitializer;
    internal readonly string SpreadsheetId;

    private static readonly string[] Scopes = { SheetsService.Scope.Drive };

    protected readonly SheetsService Service;
}
