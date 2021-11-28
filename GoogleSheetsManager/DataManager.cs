using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsManager
{
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    public static class DataManager
    {
        public static async Task<IList<T>> GetValuesAsync<T>(Provider provider, string range,
            SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum valueRenderOption =
                SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMATTEDVALUE)
            where T : ILoadable, new()
        {
            ValueRange rawValueRange = await provider.GetValuesAsync(range, valueRenderOption);
            IList<IList<object>> rawValueSets = rawValueRange.Values;
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
                    valueSet[titles[j]] = rawValueSet[j];
                }
                var instance = LoadValues<T>(valueSet);
                instances.Add(instance);
            }
            return instances;
        }

        public static Task UpdateValuesAsync<T>(Provider provider, string range, IList<T> instances)
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

            return provider.UpdateValuesAsync(range, rawValueSets);
        }

        private static T LoadValues<T>(IDictionary<string, object> valueSet) where T : ILoadable, new()
        {
            var instance = new T();
            instance.Load(valueSet);
            return instance;
        }

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

        public static bool? ToBool(this object o) => bool.TryParse(o?.ToString(), out bool b) ? (bool?)b : null;

        public static string GetHyperlink(Uri link, string text) => string.Format(HyperlinkFormat, link, text);

        private const string HyperlinkFormat = "=HYPERLINK(\"{0}\";\"{1}\")";
    }
}
