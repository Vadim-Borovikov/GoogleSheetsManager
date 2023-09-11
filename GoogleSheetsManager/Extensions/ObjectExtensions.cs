using System;
using GryphonUtilities;
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

    public static DateTime? ToDateTime(this object? o)
    {
        return o switch
        {
            double d => DateTime.FromOADate(d),
            long l   => DateTime.FromOADate(l),
            _        => null
        };
    }

    public static DateTimeFull? ToDateTimeFull(this object? o, TimeManager timeManager)
    {
        if (o is DateTimeFull dtf || DateTimeFull.TryParse(o?.ToString(), out dtf))
        {
            return dtf;
        }

        DateTimeOffset? dto = o?.ToDateTimeOffset();
        if (dto.HasValue)
        {
            return timeManager.GetDateTimeFull(dto.Value);
        }

        DateTime? dt = o.ToDateTime();
        if (dt.HasValue)
        {
            return timeManager.GetDateTimeFull(dt.Value);
        }

        return null;
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