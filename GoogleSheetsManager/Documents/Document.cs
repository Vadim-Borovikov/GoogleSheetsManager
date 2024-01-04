using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsManager.Providers;
using JetBrains.Annotations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace GoogleSheetsManager.Documents;

[PublicAPI]
public class Document
{
    public readonly string Id;

    internal readonly SheetsProvider Provider;

    public Document(SheetsService service, string id, IDictionary<Type, Func<object?, object?>> converters)
    {
        Id = id;
        Provider = new SheetsProvider(service, id);
        _converters = converters;
        _sheets = new List<Sheet>();
    }

    internal Document(SheetsService service, Spreadsheet spreadsheet,
        IDictionary<Type, Func<object?, object?>> converters)
        : this(service, spreadsheet.SpreadsheetId, converters)
    {
        _spreadsheet = spreadsheet;
    }

    public Sheet GetOrAddSheet(string title, IDictionary<Type, Func<object?, object?>>? additionalConverters = null)
    {
        Sheet? sheet = _sheets.FirstOrDefault(s => s.Name == title);
        if (sheet is null)
        {
            IDictionary<Type, Func<object?, object?>> сonverters = GetConvertersWith(additionalConverters);
            sheet = new Sheet(title, Provider, this, сonverters);
            _sheets.Add(sheet);
        }
        return sheet;
    }

    public async Task<Sheet?> GetOrAddSheetAsync(int index,
        IDictionary<Type, Func<object?, object?>>? additionalConverters = null)
    {
        Sheet? sheet = _sheets.FirstOrDefault(s => s.Index == index);
        if (sheet is null)
        {
            Spreadsheet spreadsheet = await GetSpreadsheetAsync();
            Google.Apis.Sheets.v4.Data.Sheet? googleSheet =
                spreadsheet.Sheets.SingleOrDefault(s => s.Properties.Index == index);
            if (googleSheet is null)
            {
                return null;
            }

            sheet = _sheets.FirstOrDefault(s => s.Name == googleSheet.Properties.Title);
            if (sheet is null)
            {
                IDictionary<Type, Func<object?, object?>> сonverters = GetConvertersWith(additionalConverters);
                sheet = new Sheet(googleSheet, Provider, this, сonverters);
                _sheets.Add(sheet);
            }
            else
            {
                sheet.SetSheet(googleSheet);
            }
        }

        return sheet;
    }

    internal async Task<Spreadsheet> GetSpreadsheetAsync() => _spreadsheet ??= await Provider.LoadSpreadsheetAsync();

    private IDictionary<Type, Func<object?, object?>> GetConvertersWith(
        IDictionary<Type, Func<object?, object?>>? additional)
    {
        if (additional is null)
        {
            return _converters;
        }

        IDictionary<Type, Func<object?, object?>> converters =
            new Dictionary<Type, Func<object?, object?>>(_converters);
        foreach (Type type in additional.Keys)
        {
            converters[type] = additional[type];
        }

        return converters;
    }

    private Spreadsheet? _spreadsheet;

    private readonly IList<Sheet> _sheets;
    private readonly IDictionary<Type, Func<object?, object?>> _converters;
}