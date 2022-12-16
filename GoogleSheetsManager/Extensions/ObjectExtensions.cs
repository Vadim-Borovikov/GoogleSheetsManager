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
            long l => l,
            double d => (decimal)d,
            _ => null
        };
    }

    public static DateTime? ToDateTime(this object? o)
    {
        return o switch
        {
            double d => DateTime.FromOADate(d),
            long l => DateTime.FromOADate(l),
            _ => null
        };
    }

    public static DateTimeFull? ToDateTimeFull(this object? o, TimeManager timeManager)
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
}