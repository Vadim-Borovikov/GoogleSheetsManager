using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using GryphonUtilities.Extensions;
using GryphonUtilities.Time;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Extensions;

[PublicAPI]
public static class ObjectExtensions
{
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

    public static byte? ToByte(this object? o)
    {
        if (o is byte b)
        {
            return b;
        }
        return byte.TryParse(o?.ToString(), out b) ? b : null;
    }

    public static long? ToLong(this object? o)
    {
        if (o is long l)
        {
            return l;
        }
        return long.TryParse(o?.ToString(), out l) ? l : null;
    }

    public static decimal? ToDecimal(this object? o)
    {
        return o switch
        {
            decimal d => d,
            long l    => l,
            double d  => (decimal) d,
            _         => decimal.TryParse(o?.ToString(), CultureInfo.InvariantCulture, out decimal d) ? d : null
        };
    }

    public static List<T>? ToList<T>(this object? o, string separator, Func<string, T?> converter)
    {
        if (o is IEnumerable<T> l)
        {
            return l.ToList();
        }

        string[]? parts = o?.ToString()?.Split(separator);

        return parts?.Select(converter).TryDenullAll();
    }

    public static List<T>? ToList<T>(this object? o, string separator, Func<string, T?> converter) where T : struct
    {
        if (o is IEnumerable<T> l)
        {
            return l.ToList();
        }

        string[]? parts = o?.ToString()?.Split(separator);

        return parts?.Select(converter).TryDenullAll();
    }

    public static DateOnly? ToDateOnly(this object? o, Clock clock)
    {
        if (o is DateOnly d)
        {
            return d;
        }

        DateTimeFull? dtf = o.ToDateTimeFull(clock);
        return dtf?.DateOnly;
    }

    public static TimeOnly? ToTimeOnly(this object? o, Clock clock)
    {
        if (o is TimeOnly t)
        {
            return t;
        }

        DateTimeFull? dtf = o.ToDateTimeFull(clock);
        return dtf?.TimeOnly;
    }

    public static TimeSpan? ToTimeSpan(this object? o, Clock clock)
    {
        if (o is TimeSpan t)
        {
            return t;
        }

        DateTimeFull? dtf = o.ToDateTimeFull(clock);
        return dtf?.DateTimeOffset.TimeOfDay;
    }

    public static DateTimeFull? ToDateTimeFull(this object? o, Clock clock)
    {
        if (o is DateTimeFull dtf || DateTimeFull.TryParse(o?.ToString(), out dtf))
        {
            return dtf;
        }

        DateTimeOffset? dto = o?.ToDateTimeOffset();
        if (dto.HasValue)
        {
            return clock.GetDateTimeFull(dto.Value);
        }

        DateTime? dt = o.ToDateTime();
        if (dt.HasValue)
        {
            return clock.GetDateTimeFull(dt.Value);
        }

        return null;
    }

    private static DateTime? ToDateTime(this object? o)
    {
        return o switch
        {
            double d => DateTime.FromOADate(d),
            long l   => DateTime.FromOADate(l),
            _        => null
        };
    }

    private static DateTimeOffset? ToDateTimeOffset(this object? o)
    {
        if (o is DateTimeOffset dto)
        {
            return dto;
        }
        return DateTimeOffset.TryParse(o?.ToString(), out dto) ? dto : null;
    }
}