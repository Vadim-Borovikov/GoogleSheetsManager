﻿using System;
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

        List<Dictionary<string, object?>> valueSets = rawValueSets.Select(r => Organize(r, titles)).ToList();
        return new SheetData<Dictionary<string, object?>>(valueSets, titles);
    }

    private static Dictionary<string, object?> Organize(IList<object> rawValueSet, IList<string> titles)
    {
        Dictionary<string, object?> result = new();
        for (int i = 0; i < titles.Count; ++i)
        {
            result[titles[i]] = i < rawValueSet.Count ? rawValueSet[i] : null;
        }
        return result;
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

    public static DateTimeFull? GetDateTimeFull(object? o, TimeManager timeManager)
    {
        switch (o)
        {
            case DateTimeFull dtf: return dtf;
            case DateTimeOffset dto: return timeManager.GetDateTimeFull(dto);
            default:
            {
                DateTime? dt = o.ToDateTime();
                return dt is null ? null : timeManager.GetDateTimeFull(dt.Value);
            }
        }
    }

    internal static readonly Dictionary<Type, Func<object?, object?>> DefaultConverters = new()
    {
        { typeof(bool), v => v.ToBool() },
        { typeof(bool?), v => v.ToBool() },
        { typeof(int), v => v.ToInt() },
        { typeof(int?), v => v.ToInt() },
        { typeof(decimal), v => v.ToDecimal() },
        { typeof(decimal?), v => v.ToDecimal() },
        { typeof(string), v => v?.ToString() },
    };

    public static string GetHyperlink(Uri uri, string? caption = null)
    {
        if (string.IsNullOrWhiteSpace(caption))
        {
            caption = uri.AbsoluteUri;
        }
        return string.Format(HyperlinkFormat, uri.AbsoluteUri, caption);
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

    private const string HyperlinkFormat = "=HYPERLINK(\"{0}\";\"{1}\")";
}