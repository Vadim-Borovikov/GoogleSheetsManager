using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;
using GoogleSheetsManager.Providers;

namespace GoogleSheetsManager
{
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    public static class DataManager
    {
        public static async Task<IList<T>> GetValuesAsync<T>(SheetsProvider provider, string range, bool formula = false)
            where T : ILoadable, new()
        {
            IList<IList<object>> rawValueSets = await provider.GetValueListAsync(range, formula);
            if (rawValueSets.Count < 1)
            {
                return new List<T>();
            }
            List<string> titles = rawValueSets[0].Select(o => o.ToString()).ToList();
            var instances = new List<T>();
            for (int i = 1; i < rawValueSets.Count; ++i)
            {
                var valueSet = new Dictionary<string, object>();
                IList<object> rawValueSet = rawValueSets[i];
                for (int j = 0; j < titles.Count; ++j)
                {
                    string title = titles[j];
                    valueSet[title] = j < rawValueSet.Count ? rawValueSet[j] : null;
                }
                var instance = LoadValues<T>(valueSet);
                instances.Add(instance);
            }
            return instances;
        }

        public static Task UpdateValuesAsync<T>(SheetsProvider sheetsProvider, string range, IList<T> instances)
            where T : ISavable
        {
            IList<string> titles = instances[0].Titles;
            var rawValueSets = new List<IList<object>> { titles.Cast<object>().ToList() };

            IEnumerable<IDictionary<string, object>> valueSets = instances.Select(v => v.Save());
            // ReSharper disable once LoopCanBeConvertedToQuery
            foreach (IDictionary<string, object> valueSet in valueSets)
            {
                List<object> rawValueSet = titles.Select(t => valueSet[t]).ToList();
                rawValueSets.Add(rawValueSet);
            }

            return sheetsProvider.UpdateValuesAsync(range, rawValueSets);
        }

        public static async Task<string> CopyForAsync(SheetsProvider sheetsProvider, string name, string folderId,
            string ownerEmail)
        {
            await sheetsProvider.LoadSpreadsheetAsync();

            using (SheetsProvider newSheetsProvider = await sheetsProvider.CreateNewWithPropertiesAsync())
            {
                await CopyContentAsync(sheetsProvider, newSheetsProvider);

                using (var driveProvider =
                    new DriveProvider(sheetsProvider.ServiceInitializer, newSheetsProvider.SpreadsheetId))
                {
                    await SetupPermissionsForAsync(driveProvider, ownerEmail);

                    await MoveAndRenameAsync(driveProvider, name, folderId);
                }

                return newSheetsProvider.SpreadsheetId;
            }
        }

        public static TimeSpan? ToTimeSpan(this object o) => ToDateTime(o)?.TimeOfDay;

        public static Uri ToUri(this object o)
        {
            string uriString = o?.ToString();
            // ReSharper disable once AssignNullToNotNullAttribute
            return string.IsNullOrWhiteSpace(uriString) ? null : new Uri(uriString);
        }

        public static decimal? ToDecimal(this object o)
        {
            switch (o)
            {
                case long l:
                    return l;
                case double d:
                    return (decimal)d;
                default:
                    return null;
            }
        }

        public static int? ToInt(this object o) => int.TryParse(o?.ToString(), out int i) ? (int?)i : null;

        public static long? ToLong(this object o) => long.TryParse(o?.ToString(), out long l) ? (long?)l : null;

        public static ushort? ToUshort(this object o) => ushort.TryParse(o?.ToString(), out ushort u) ? (ushort?)u : null;

        public static byte? ToByte(this object o) => byte.TryParse(o?.ToString(), out byte b) ? (byte?)b : null;

        public static bool? ToBool(this object o) => bool.TryParse(o?.ToString(), out bool b) ? (bool?)b : null;

        public static List<Uri> ToUris(this object o) => o?.ToString()
                                                          ?.Split("\n")
                                                          ?.Select(ToUri)
                                                          .Where(u => u != null)
                                                          .ToList();

        public static DateTime? ToDateTime(this object o)
        {
            switch (o)
            {
                case double d:
                    return DateTime.FromOADate(d);
                case long l:
                    return DateTime.FromOADate(l);
                default:
                    return null;
            }
        }

        public static string GetHyperlink(Uri link, string text) => string.Format(HyperlinkFormat, link, text);

        private static T LoadValues<T>(IDictionary<string, object> valueSet) where T : ILoadable, new()
        {
            var instance = new T();
            instance.Load(valueSet);
            return instance;
        }

        private static async Task CopyContentAsync(SheetsProvider from, SheetsProvider to)
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

        private static Task AddPermissionToAsync(DriveProvider provider, string role, string type, string emailAddress = null)
        {
            bool transferOwnership = role == "owner";
            return provider.AddPermissionToAsync(type, role, emailAddress, transferOwnership);
        }

        private static async Task MoveAndRenameAsync(DriveProvider provider, string name, string folderId)
        {
            IList<string> oldParentLists = await provider.GetParentsAsync();
            string oldParents = string.Join(',', oldParentLists);
            await provider.MoveAndRenameAsync(name, folderId, oldParents);
        }

        private const string HyperlinkFormat = "=HYPERLINK(\"{0}\";\"{1}\")";
    }
}
