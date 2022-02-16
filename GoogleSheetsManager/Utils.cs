using System;
using System.Collections.Generic;
using System.Linq;
using JetBrains.Annotations;

namespace GoogleSheetsManager;

[PublicAPI]
public static class Utils
{
    public static TimeSpan? ToTimeSpan(this object? o) => ToDateTime(o)?.TimeOfDay;

    public static Uri? ToUri(this object? o)
    {
        string? uriString = o?.ToString();
        return string.IsNullOrWhiteSpace(uriString) ? null : new Uri(uriString);
    }

    public static decimal? ToDecimal(this object? o)
    {
        return o switch
        {
            long l => l,
            double d => (decimal)d,
            _ => null
        };
    }

    public static int? ToInt(this object? o) => int.TryParse(o?.ToString(), out int i) ? i : null;

    public static long? ToLong(this object? o) => long.TryParse(o?.ToString(), out long l) ? l : null;

    public static ushort? ToUshort(this object? o) => ushort.TryParse(o?.ToString(), out ushort u) ? u : null;

    public static byte? ToByte(this object? o) => byte.TryParse(o?.ToString(), out byte b) ? b : null;

    public static bool? ToBool(this object? o) => bool.TryParse(o?.ToString(), out bool b) ? b : null;

    public static List<Uri>? ToUris(this object? o) => o?.ToString()?.Split("\n").Select(ToUri).RemoveNulls().ToList();

    public static DateTime? ToDateTime(this object? o)
    {
        return o switch
        {
            double d => DateTime.FromOADate(d),
            long l   => DateTime.FromOADate(l),
            _        => null
        };
    }

    public static IEnumerable<T> RemoveNulls<T>(this IEnumerable<T?> seq)
    {
        return seq.Where(i => i is not null).Select(i => i.GetValue());
    }
    public static IEnumerable<T> RemoveNulls<T>(this IEnumerable<T?> seq) where T : struct
    {
        return seq.Where(i => i is not null).Select(i => i.GetValue());
    }

    public static T GetValue<T>(this T? param, string? message = null)
    {
        return param ?? throw new NullReferenceException(message);
    }
    public static T GetValue<T>(this T? param, string? message = null) where T : struct
    {
        return param ?? throw new NullReferenceException(message);
    }
}
