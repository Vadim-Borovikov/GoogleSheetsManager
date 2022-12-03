using System;
using System.Collections.Generic;
using System.Linq;
using GoogleSheetsManager.Providers;
using System.Threading.Tasks;
using JetBrains.Annotations;
using Google.Apis.Sheets.v4.Data;
using GryphonUtilities;

namespace GoogleSheetsManager;

[PublicAPI]
public static class Utils
{
    public static async Task<SheetData<Dictionary<string, object?>>> LoadAsync(SheetsProvider provider, string range,
        int? sheetIndex = null, bool formula = false)
    {
        if (sheetIndex.HasValue)
        {
            Sheet sheet = await GetSheetAsync(provider, sheetIndex.Value);
            range = $"{sheet.Properties.Title}!{range}";
        }

        IList<IList<object>> rawValueSets = await provider.GetValueListAsync(range, formula);
        if (rawValueSets.Count < 1)
        {
            return new SheetData<Dictionary<string, object?>>();
        }
        List<string> titles = GetTitles(rawValueSets[0]).ToList();

        List<Dictionary<string, object?>> valueSets = new();
        for (int i = 1; i < rawValueSets.Count; ++i)
        {
            Dictionary<string, object?> valueSet = new();
            IList<object> rawValueSet = rawValueSets[i];
            for (int j = 0; j < titles.Count; ++j)
            {
                string title = titles[j];
                valueSet[title] = j < rawValueSet.Count ? rawValueSet[j] : null;
            }
            valueSets.Add(valueSet);
        }

        return new SheetData<Dictionary<string, object?>>(valueSets, titles);
    }

    public static async Task<List<string>> LoadTitlesAsync(SheetsProvider provider, string range)
    {
        range = GetFirstRow(range);
        IList<IList<object>> rawValueSets = await provider.GetValueListAsync(range, false);
        return GetTitles(rawValueSets[0]).ToList();
    }

    public static Task SaveAsync(SheetsProvider provider, string range, SheetData<Dictionary<string, object?>> data)
    {
        List<IList<object>> rawValueSets = new() { data.Titles.ToList<object>() };
        rawValueSets.AddRange(data.Instances
                                  .Select(set => data.Titles.Select(t => set.TryGetValue(t) ?? "").ToList()));
        return provider.UpdateValuesAsync(range, rawValueSets);
    }

    public static async Task<string> CopyForAsync(SheetsProvider from, string name, string folderId)
    {
        Spreadsheet fromSpreadsheet = await from.LoadSpreadsheetAsync();
        Spreadsheet toSpreadsheet = await from.CreateNewSpreadsheetAsync(fromSpreadsheet.Properties);
        using (SheetsProvider to =
               new(from.ServiceInitializer, from.Service, from.TimeManager, toSpreadsheet.SpreadsheetId))
        {
            to.PlanToDeleteSheets(toSpreadsheet);
            await to.CopyContentAndPlanToRenameSheetsAsync(from, fromSpreadsheet);
            await to.ExecutePlannedAsync();

            using (DriveProvider driveProvider = new(from.ServiceInitializer, to.SpreadsheetId))
            {
                await driveProvider.AddPermissionToAsync("anyone", "writer", null);
                await MoveAndRenameAsync(driveProvider, name, folderId);
            }

            return to.SpreadsheetId;
        }
    }

    public static async Task DeleteSpreadsheetAsync(SheetsProvider provider, string fileId)
    {
        using (DriveProvider driveProvider = new(provider.ServiceInitializer, fileId))
        {
            await driveProvider.DeleteSpreadsheetAsync();
        }
    }

    public static async Task RenameSheetAsync(SheetsProvider provider, int sheetIndex, string title)
    {
        Sheet sheet = await GetSheetAsync(provider, sheetIndex);
        await provider.RenameSheetAsync(sheet.Properties.SheetId, title);
    }

    public static string GetHyperlink(Uri uri, string? caption = null)
    {
        if (string.IsNullOrWhiteSpace(caption))
        {
            caption = uri.AbsoluteUri;
        }
        return string.Format(HyperlinkFormat, uri.AbsoluteUri, caption);
    }

    public static DateTimeFull? GetDateTimeFull(object? o, TimeManager timeManager)
    {
        switch (o)
        {
            case DateTimeFull dtf: return dtf;
            case DateTimeOffset dto: return timeManager.GetDateTimeFull(dto);
            default:
            {
                DateTime? dt = GetDateTime(o);
                return dt is null ? null : timeManager.GetDateTimeFull(dt.Value);
            }
        }
    }

    public static bool? ToBool(this object? o)
    {
        if (o is bool b)
        {
            return b;
        }
        return bool.TryParse(o?.ToString(), out b) ? b : null;
    }

    public static int? ToInt(this object? o)
    {
        if (o is int i)
        {
            return i;
        }
        return int.TryParse(o?.ToString(), out i) ? i : null;
    }

    public static decimal? ToDecimal(this object? o)
    {
        return o switch
        {
            decimal dec => dec,
            long l      => l,
            double d    => (decimal) d,
            _           => null
        };
    }

    public static DateTime? GetDateTime(object? o)
    {
        return o switch
        {
            double d => DateTime.FromOADate(d),
            long l   => DateTime.FromOADate(l),
            _        => null
        };
    }

    private static IEnumerable<string> GetTitles(IEnumerable<object> rawValueSet)
    {
        return rawValueSet.Select(o => o.ToString() ?? "");
    }

    private static async Task MoveAndRenameAsync(DriveProvider provider, string name, string folderId)
    {
        IList<string> oldParentLists = await provider.GetParentsAsync();
        string oldParents = string.Join(',', oldParentLists);
        await provider.MoveAndRenameAsync(name, folderId, oldParents);
    }

    private static async Task<Sheet> GetSheetAsync(SheetsProvider provider, int sheetIndex)
    {
        Spreadsheet spreadsheet = await provider.LoadSpreadsheetAsync();
        return spreadsheet.Sheets[sheetIndex];
    }

    private static string GetFirstRow(string range)
    {
        Range.Range? r = Range.Range.Parse(range);
        if (r is null)
        {
            return range;
        }

        if (r.IntervalEnd is null)
        {
            return range;
        }

        r.IntervalEnd.Row = r.IntervalStart.Row;
        return r.ToString();
    }

    private static object? TryGetValue(this Dictionary<string, object?> set, string key)
    {
        set.TryGetValue(key, out object? o);
        return o;
    }

    internal static readonly Dictionary<Type, Func<object?, object?>> DefaultConverters = new()
    {
        { typeof(bool), v => ToBool(v) },
        { typeof(bool?), v => ToBool(v) },
        { typeof(int), v => ToInt(v) },
        { typeof(int?), v => ToInt(v) },
        { typeof(decimal), v => ToDecimal(v) },
        { typeof(decimal?), v => ToDecimal(v) },
        { typeof(string), v => v?.ToString() },
    };

    private const string HyperlinkFormat = "=HYPERLINK(\"{0}\";\"{1}\")";
}