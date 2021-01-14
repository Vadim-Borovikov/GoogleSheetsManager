using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace GoogleSheetsManager
{
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    public static class DataManager
    {
        public static IList<T> GetValues<T>(Provider provider, string range) where T : ILoadable, new()
        {
            IEnumerable<IList<object>> values = provider.GetValues(range, true);
            return values?.Select(LoadValues<T>).ToList();
        }

        public static void UpdateValues<T>(Provider provider, string range, IEnumerable<T> values)
            where T : ISavable
        {
            List<IList<object>> table = values.Select(v => v.Save()).ToList();
            provider.UpdateValues(range, table);
        }

        public static string ToString(this IList<object> values, int index) => To(values, index, o => o?.ToString());

        public static DateTime? ToDateTime(this IList<object> values, int index) => To(values, index, ToDateTime);
        public static TimeSpan? ToTimeSpan(IList<object> values, int index) => To(values, index, ToTimeSpan);
        public static Uri ToUri(IList<object> values, int index) => To(values, index, ToUri);
        public static decimal? ToDecimal(this IList<object> values, int index) => To(values, index, ToDecimal);
        public static int? ToInt(this IList<object> values, int index) => To(values, index, ToInt);

        public static T To<T>(this IList<object> values, int index, Func<object, T> cast)
        {
            object o = values.Count > index ? values[index] : null;
            return cast(o);
        }

        private static T LoadValues<T>(IList<object> values) where T : ILoadable, new()
        {
            var instance = new T();
            instance.Load(values);
            return instance;
        }

        private static DateTime? ToDateTime(object o)
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

        private static TimeSpan? ToTimeSpan(object o) => ToDateTime(o)?.TimeOfDay;

        private static Uri ToUri(object o)
        {
            string uriString = o?.ToString();
            return string.IsNullOrWhiteSpace(uriString) ? null : new Uri(uriString);
        }

        private static decimal? ToDecimal(object o)
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

        public static string GetHyperlink(Uri link, string text) => string.Format(HyperlinkFormat, link, text);

        private static int? ToInt(object o) => int.TryParse(o?.ToString(), out int i) ? (int?)i : null;

        private const string HyperlinkFormat = "=HYPERLINK(\"{0}\";\"{1}\")";
    }
}
