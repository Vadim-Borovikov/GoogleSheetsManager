using System;
using System.Collections.Generic;
using System.Linq;
using GryphonUtilities;
using JetBrains.Annotations;

namespace GoogleSheetsManager;

[PublicAPI]
public static class Utils
{
    public static TimeSpan? ToTimeSpan(this object? o)
    {
        if (o is TimeSpan ts)
        {
            return ts;
        }
        return ToDateTime(o)?.TimeOfDay;
    }

    public static Uri? ToUri(this object? o)
    {
        if (o is Uri uri)
        {
            return uri;
        }
        string? uriString = o?.ToString();
        return string.IsNullOrWhiteSpace(uriString) ? null : new Uri(uriString);
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

    public static int? ToInt(this object? o)
    {
        if (o is int i)
        {
            return i;
        }
        return int.TryParse(o?.ToString(), out i) ? i : null;
    }

    public static long? ToLong(this object? o)
    {
        if (o is long l)
        {
            return l;
        }
        return long.TryParse(o?.ToString(), out l) ? l : null;
    }

    public static ushort? ToUshort(this object? o)
    {
        if (o is ushort u)
        {
            return u;
        }
        return ushort.TryParse(o?.ToString(), out u) ? u : null;
    }

    public static byte? ToByte(this object? o)
    {
        if (o is byte b)
        {
            return b;
        }
        return byte.TryParse(o?.ToString(), out b) ? b : null;
    }

    public static bool? ToBool(this object? o)
    {
        if (o is bool b)
        {
            return b;
        }
        return bool.TryParse(o?.ToString(), out b) ? b : null;
    }

    public static List<Uri>? ToUris(this object? o)
    {
        if (o is IEnumerable<Uri> l)
        {
            return l.ToList();
        }
        return o?.ToString()?.Split("\n").Select(ToUri).RemoveNulls().ToList();
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
}