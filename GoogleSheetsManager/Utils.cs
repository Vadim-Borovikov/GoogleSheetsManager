﻿using System;
using System.Collections.Generic;
using System.Linq;
using GoogleSheetsManager.Providers;
using System.Threading.Tasks;
using JetBrains.Annotations;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsManager;

[PublicAPI]
public static class Utils
{
    public static async Task<SheetData<Dictionary<string, object?>>> LoadAsync(SheetsProvider provider, string range,
        int? sheetIndex = null, bool formula = false)
    {
        if (sheetIndex.HasValue)
        {
            Sheet sheet = await GetSheet(provider, sheetIndex.Value);
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
        List<IList<object>> rawValueSets = new() { data.Titles.Cast<object>().ToList() };
        rawValueSets.AddRange(data.Instances
                                  .Select(set => data.Titles.Select(t => set[t] ?? "").ToList()));
        return provider.UpdateValuesAsync(range, rawValueSets);
    }

    public static async Task<string> CopyForAsync(SheetsProviderWithSpreadsheet sheetsProvider, string name,
        string folderId, string ownerEmail)
    {
        using (SheetsProviderWithSpreadsheet newSheetsProvider = await sheetsProvider.CreateNewWithPropertiesAsync())
        {
            await CopyContentAsync(sheetsProvider, newSheetsProvider);

            using (DriveProvider driveProvider =
                   new(sheetsProvider.ServiceInitializer, newSheetsProvider.SpreadsheetId))
            {
                await SetupPermissionsForAsync(driveProvider, ownerEmail);

                await MoveAndRenameAsync(driveProvider, name, folderId);
            }

            return newSheetsProvider.SpreadsheetId;
        }
    }

    public static async Task RenameSheetAsync(SheetsProvider provider, int sheetIndex, string title)
    {
        Sheet sheet = await GetSheet(provider, sheetIndex);
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
            double d    => (decimal)d,
            _           => null
        };
    }

    public static DateTime? ToDateTime(this object? o)
    {
        return o switch
        {
            DateTime dt => dt,
            double d    => DateTime.FromOADate(d),
            long l      => DateTime.FromOADate(l),
            _           => null
        };
    }

    public static TimeSpan? ToTimeSpan(this object? o)
    {
        if (o is TimeSpan ts)
        {
            return ts;
        }
        return ToDateTime(o)?.TimeOfDay;
    }

    private static IEnumerable<string> GetTitles(IEnumerable<object> rawValueSet)
    {
        return rawValueSet.Select(o => o.ToString() ?? "");
    }

    private static async Task CopyContentAsync(SheetsProviderWithSpreadsheet from, SheetsProviderWithSpreadsheet to)
    {
        to.PlanToDeleteSheets();
        await to.CopyContentAndPlanToRenameSheetsAsync(from);
        await to.ExecutePlanned();
    }

    private static async Task SetupPermissionsForAsync(DriveProvider provider, string ownerEmail)
    {
        IEnumerable<string> oldPermissionIds = await provider.GetPermissionIdsAsync();

        await AddPermissionToAsync(provider, "owner", "user", ownerEmail);

        await AddPermissionToAsync(provider, "writer", "anyone");

        foreach (string id in oldPermissionIds)
        {
            await provider.DowngradePermissionAsync(id);
        }
    }

    private static Task AddPermissionToAsync(DriveProvider provider, string role, string type,
        string? emailAddress = null)
    {
        bool transferOwnership = role is "owner";
        return provider.AddPermissionToAsync(type, role, emailAddress, transferOwnership);
    }

    private static async Task MoveAndRenameAsync(DriveProvider provider, string name, string folderId)
    {
        IList<string> oldParentLists = await provider.GetParentsAsync();
        string oldParents = string.Join(',', oldParentLists);
        await provider.MoveAndRenameAsync(name, folderId, oldParents);
    }

    private static async Task<Sheet> GetSheet(SheetsProvider provider, int sheetIndex)
    {
        Spreadsheet spreadsheet = await provider.LoadSpreadsheet();
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

    internal static readonly Dictionary<Type, Func<object?, object?>> DefaultConverters = new()
    {
        { typeof(bool), v => ToBool(v) },
        { typeof(bool?), v => ToBool(v) },
        { typeof(int), v => ToInt(v) },
        { typeof(int?), v => ToInt(v) },
        { typeof(decimal), v => ToDecimal(v) },
        { typeof(decimal?), v => ToDecimal(v) },
        { typeof(string), v => v?.ToString() },
        { typeof(DateTime), v => ToDateTime(v) },
        { typeof(DateTime?), v => ToDateTime(v) },
        { typeof(TimeSpan), v => ToTimeSpan(v) },
        { typeof(TimeSpan?), v => ToTimeSpan(v) },
    };

    private const string HyperlinkFormat = "=HYPERLINK(\"{0}\";\"{1}\")";
}