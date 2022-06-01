using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsManager.Providers;
using GryphonUtilities;
using JetBrains.Annotations;

namespace GoogleSheetsManager;

[PublicAPI]
public static class DataManager
{
    public static async Task<IList<T>> GetValuesAsync<T>(SheetsProvider provider,
        Func<IDictionary<string, object?>, T?> loader, string range, bool formula = false)
    {
        IList<IList<object>> rawValueSets = await provider.GetValueListAsync(range, formula);
        if (rawValueSets.Count < 1)
        {
            return Array.Empty<T>();
        }
        List<string> titles = rawValueSets[0].Select(o => o.ToString() ?? "").ToList();
        List<T> instances = new();
        for (int i = 1; i < rawValueSets.Count; ++i)
        {
            Dictionary<string, object?> valueSet = new();
            IList<object> rawValueSet = rawValueSets[i];
            for (int j = 0; j < titles.Count; ++j)
            {
                string title = titles[j];
                valueSet[title] = j < rawValueSet.Count ? rawValueSet[j] : null;
            }
            T? instance = loader(valueSet);
            if (instance is not null)
            {
                instances.Add(instance);
            }
        }
        return instances;
    }

    public static Task UpdateValuesAsync<T>(SheetsProvider sheetsProvider, string range, IList<T> instances)
        where T : ISavable
    {
        IList<string> titles = instances[0].Titles;
        List<IList<object>> rawValueSets = new() { titles.Cast<object>().ToList() };
        rawValueSets.AddRange(instances.Select(i => i.Convert())
                                       .RemoveNulls()
                                       .Select(set => titles.Select(t => set[t] ?? "").ToList()));
        return sheetsProvider.UpdateValuesAsync(range, rawValueSets);
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

    public static async Task<IList<T>> GetValuesAsync<T>(SheetsProvider provider,
        Func<IDictionary<string, object?>, T?> loader, int sheetIndex, string range, bool formula = false)
    {
        Sheet sheet = await GetSheet(provider, sheetIndex);
        range = $"{sheet.Properties.Title}!{range}";
        return await GetValuesAsync(provider, loader, range, formula);
    }

    public static string GetHyperlink(Uri link, string text) => string.Format(HyperlinkFormat, link, text);

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

    private const string HyperlinkFormat = "=HYPERLINK(\"{0}\";\"{1}\")";
}
