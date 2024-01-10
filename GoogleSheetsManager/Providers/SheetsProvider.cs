using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GryphonUtilities.Extensions;

namespace GoogleSheetsManager.Providers;

internal sealed class SheetsProvider
{
    public SheetsProvider(SheetsService service, string spreadsheetId)
    {
        _service = service;
        _spreadsheetId = spreadsheetId;
    }

    internal Task ClearValuesAsync(string range)
    {
        SpreadsheetsResource.ValuesResource.ClearRequest request =
            _service.Spreadsheets.Values.Clear(new ClearValuesRequest(), _spreadsheetId, range);
        return request.ExecuteAsync();
    }

    public async Task<IList<IList<object>>> GetValueListAsync(string range, bool formula)
    {
        SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum renderOption = formula
            ? SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMULA
            : SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.UNFORMATTEDVALUE;
        ValueRange valueRange = await GetValuesAsync(range, renderOption);
        return valueRange.Values;
    }

    public Task AppendValuesAsync(string range, IList<IList<object>> values)
    {
        ValueRange body = new() { Values = values };
        SpreadsheetsResource.ValuesResource.AppendRequest request =
            _service.Spreadsheets.Values.Append(body, _spreadsheetId, range);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        return request.ExecuteAsync();
    }

    public Task UpdateValuesAsync(string range, IList<IList<object>> values)
    {
        ValueRange body = new() { Values = values };
        SpreadsheetsResource.ValuesResource.UpdateRequest request =
            _service.Spreadsheets.Values.Update(body, _spreadsheetId, range);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
        return request.ExecuteAsync();
    }

    public Task<Spreadsheet> LoadSpreadsheetAsync()
    {
        SpreadsheetsResource.GetRequest request = _service.Spreadsheets.Get(_spreadsheetId);
        request.IncludeGridData = true;
        return request.ExecuteAsync();
    }

    public async Task CopyContentAndPlanToRenameSheetsAsync(SheetsProvider from, Spreadsheet spreadsheet)
    {
        foreach (Sheet sheet in spreadsheet.Sheets.Where<Sheet>(s => s.Properties.SheetId.HasValue))
        {
            // ReSharper disable once NullableWarningSuppressionIsUsed
            //   sheet.Properties.SheetId is null-checked already
            SheetProperties properties = await from.CopyToAsync(_spreadsheetId, sheet.Properties.SheetId!.Value);
            Request renameRequest = CreateRenameSheetRequest(properties.SheetId, sheet.Properties.Title);
            _requestsToExecute.Add(renameRequest);
        }
    }

    public Task ExecutePlannedAsync()
    {
        BatchUpdateSpreadsheetRequest body = new() { Requests = _requestsToExecute };
        SpreadsheetsResource.BatchUpdateRequest request = new(_service, body, _spreadsheetId);
        return request.ExecuteAsync();
    }

    private Task<SheetProperties> CopyToAsync(string destinationSpreadsheetId, int sheetId)
    {
        CopySheetToAnotherSpreadsheetRequest body = new() { DestinationSpreadsheetId = destinationSpreadsheetId };
        SpreadsheetsResource.SheetsResource.CopyToRequest request = new(_service, body, _spreadsheetId, sheetId);
        return request.ExecuteAsync();
    }

    public Task RenameSheetAsync(int? sheetId, string title)
    {
        Request request = CreateRenameSheetRequest(sheetId, title);
        BatchUpdateSpreadsheetRequest body = new() { Requests = request.WrapWithList() };
        SpreadsheetsResource.BatchUpdateRequest batchRequest = _service.Spreadsheets.BatchUpdate(body, _spreadsheetId);
        return batchRequest.ExecuteAsync();
    }

    public Task<Spreadsheet> CreateNewSpreadsheetAsync(SpreadsheetProperties properties)
    {
        Spreadsheet body = new() { Properties = properties };
        SpreadsheetsResource.CreateRequest request = new(_service, body);
        return request.ExecuteAsync();
    }

    public void PlanToDeleteSheets(Spreadsheet spreadsheet)
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
            _service.Spreadsheets.Values.Get(_spreadsheetId, range);
        request.ValueRenderOption = valueRenderOption;
        request.DateTimeRenderOption =
            SpreadsheetsResource.ValuesResource.GetRequest.DateTimeRenderOptionEnum.SERIALNUMBER;
        return request.ExecuteAsync();
    }

    internal static readonly string[] Scopes = { SheetsService.Scope.Drive };

    private readonly SheetsService _service;
    private readonly string _spreadsheetId;
}
