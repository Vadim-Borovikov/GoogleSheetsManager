using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Providers;

[PublicAPI]
public sealed class SheetsProviderWithSpreadsheet : SheetsProvider
{
    public static async Task<SheetsProviderWithSpreadsheet> CreateAsync(string credentialJson, string applicationName,
        string spreadsheetId)
    {
        BaseClientService.Initializer initializer = CreateInitializer(credentialJson, applicationName);
        SheetsService service = new(initializer);
        SpreadsheetsResource.GetRequest request = new(service, spreadsheetId);
        Spreadsheet spreadsheet = await request.ExecuteAsync();
        return new SheetsProviderWithSpreadsheet(initializer, service, spreadsheet);
    }

    private SheetsProviderWithSpreadsheet(BaseClientService.Initializer initializer, SheetsService service,
        Spreadsheet spreadsheet)
        : base(initializer, service, spreadsheet.SpreadsheetId)
    {
        _spreadsheet = spreadsheet;
    }

    internal async Task<SheetsProviderWithSpreadsheet> CreateNewWithPropertiesAsync()
    {
        Spreadsheet spreadsheet = await CreateNewAsync(_spreadsheet.Properties);
        return new SheetsProviderWithSpreadsheet(ServiceInitializer, Service, spreadsheet);
    }

    internal void PlanToDeleteSheets()
    {
        _requestsToExecute.Clear();
        _requestsToExecute.AddRange(_spreadsheet.Sheets.Select(s => CreateDeleteSheetRequest(s.Properties.SheetId)));
    }

    internal async Task CopyContentAndPlanToRenameSheetsAsync(SheetsProviderWithSpreadsheet from)
    {
        foreach (Sheet sheet in from._spreadsheet.Sheets.Where(s => s.Properties.SheetId.HasValue))
        {
            // ReSharper disable once NullableWarningSuppressionIsUsed
            //   sheet.Properties.SheetId is null-checked already
            SheetProperties properties = await from.CopyToAsync(SpreadsheetId, sheet.Properties.SheetId!.Value);
            Request renameRequest = CreateRenameSheetRequest(properties.SheetId, sheet.Properties.Title);
            _requestsToExecute.Add(renameRequest);
        }
    }

    internal Task ExecutePlanned() => BatchUpdateAsync(_requestsToExecute);

    private Task<SheetProperties> CopyToAsync(string destinationSpreadsheetId, int sheetId)
    {
        CopySheetToAnotherSpreadsheetRequest body = new() { DestinationSpreadsheetId = destinationSpreadsheetId };
        SpreadsheetsResource.SheetsResource.CopyToRequest request = new(Service, body, SpreadsheetId, sheetId);
        return request.ExecuteAsync();
    }

    private Task BatchUpdateAsync(IList<Request> requests)
    {
        BatchUpdateSpreadsheetRequest body = new() { Requests = requests };
        SpreadsheetsResource.BatchUpdateRequest request = new(Service, body, SpreadsheetId);
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

    private Task<Spreadsheet> CreateNewAsync(SpreadsheetProperties properties)
    {
        Spreadsheet body = new() { Properties = properties };
        SpreadsheetsResource.CreateRequest request = new(Service, body);
        return request.ExecuteAsync();
    }

    private readonly Spreadsheet _spreadsheet;
    private readonly List<Request> _requestsToExecute = new();
}
