using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsManager.Providers;
using JetBrains.Annotations;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GoogleSheetsManager.Documents;

[PublicAPI]
public class Document
{
    internal readonly SheetsProvider Provider;

    public Document(SheetsService service, string id, IDictionary<Type, Func<object?, object?>> converters)
    {
        Provider = new SheetsProvider(service, id);
        _converters = converters;
        _sheets = new Dictionary<string, Sheet>();
    }

    internal Document(SheetsService service, Spreadsheet spreadsheet,
        IDictionary<Type, Func<object?, object?>> converters)
        : this(service, spreadsheet.SpreadsheetId, converters)
    {
        _spreadsheet = spreadsheet;
    }

    public Sheet GetOrAddSheet(string title, IDictionary<Type, Func<object?, object?>>? additionalConverters = null)
    {
        if (!_sheets.ContainsKey(title))
        {
            _sheets[title] = AddSheet(title, additionalConverters);
        }

        return _sheets[title];
    }

    internal async Task<Spreadsheet> GetSpreadsheetAsync() => _spreadsheet ??= await Provider.LoadSpreadsheetAsync();

    private Sheet AddSheet(string title, IDictionary<Type, Func<object?, object?>>? additionalConverters)
    {
        IDictionary<Type, Func<object?, object?>> converters;
        if (additionalConverters is null)
        {
            converters = _converters;
        }
        else
        {
            converters = new Dictionary<Type, Func<object?, object?>>(_converters);
            foreach (Type type in additionalConverters.Keys)
            {
                converters[type] = additionalConverters[type];
            }
        }

        return new Sheet(title, Provider, this, converters);
    }

    private Spreadsheet? _spreadsheet;

    private readonly IDictionary<string, Sheet> _sheets;
    private readonly IDictionary<Type, Func<object?, object?>> _converters;
}