using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using JetBrains.Annotations;

namespace GoogleSheetsManager;

[PublicAPI]
public static class Utils
{
    public static IDictionary<string, object?> Save<T>(this T instance)
        where T : class
    {
        Dictionary<string, object?> result = new();
        Type type = typeof(T);
        instance.Save(result, type.GetProperties(), (info, obj) => info.GetValue(obj));
        instance.Save(result, type.GetFields(), (info, obj) => info.GetValue(obj));
        return result;
    }

    public static T? Load<T>(IDictionary<string, object?> valueSet)
        where T : class, new()
    {
        T? result = new();
        Type type = typeof(T);
        result = Load(valueSet, result, type.GetProperties(), info => info.PropertyType,
            (info, obj, val) => info.SetValue(obj, val));
        result = Load(valueSet, result, type.GetFields(), info => info.FieldType,
            (info, obj, val) => info.SetValue(obj, val));
        return result;
    }

    internal static string GetFirstRow(string range)
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

    private static void Save<TInstance, TInfo>(this TInstance instance, Dictionary<string, object?> result,
        IEnumerable<TInfo> members, Func<TInfo, TInstance, object?> getter)
        where TInstance : class
        where TInfo : MemberInfo
    {
        foreach (TInfo info in members)
        {
            SheetFieldAttribute? sheetField = info.GetCustomAttributes<SheetFieldAttribute>(true).SingleOrDefault();
            if (sheetField is null)
            {
                continue;
            }

            object? value = getter(info, instance);
            result[sheetField.Title] = sheetField.Format is null ? value : string.Format(sheetField.Format, value);
        }
    }

    private static TInstance? Load<TInstance, TInfo>(IDictionary<string, object?> valueSet, TInstance? result,
        IEnumerable<TInfo> members, Func<TInfo, Type> typeProvider, Action<TInfo, TInstance, object?> setter)
        where TInstance : class
        where TInfo : MemberInfo
    {
        if (result is null)
        {
            return null;
        }

        foreach (TInfo info in members)
        {
            bool required = false;
            string? title = null;
            foreach (Attribute attribute in info.GetCustomAttributes())
            {
                switch (attribute)
                {
                    case RequiredAttribute: required = true;
                        break;
                    case SheetFieldAttribute sheetField: title = sheetField.Title;
                        break;
                }
            }

            if (title is null)
            {
                continue;
            }

            if (required && !valueSet.ContainsKey(title))
            {
                return null;
            }

            Type type = typeProvider(info);
            Func<object?, object?>? converter = Converters.GetValueOrDefault(type);
            object? value = converter?.Invoke(valueSet[title]);
            if (required && value is null or "")
            {
                return null;
            }
            setter(info, result, value);
        }

        return result;
    }

    private static bool? ToBool(this object? o)
    {
        if (o is bool b)
        {
            return b;
        }
        return bool.TryParse(o?.ToString(), out b) ? b : null;
    }

    private static int? ToInt(this object? o)
    {
        if (o is int i)
        {
            return i;
        }
        return int.TryParse(o?.ToString(), out i) ? i : null;
    }

    private static decimal? ToDecimal(this object? o)
    {
        return o switch
        {
            decimal dec => dec,
            long l      => l,
            double d    => (decimal) d,
            _           => null
        };
    }

    private static DateTime? ToDateTime(this object? o)
    {
        return o switch
        {
            DateTime dt => dt,
            double d    => DateTime.FromOADate(d),
            long l      => DateTime.FromOADate(l),
            _           => null
        };
    }

    private static TimeSpan? ToTimeSpan(this object? o)
    {
        if (o is TimeSpan ts)
        {
            return ts;
        }
        return ToDateTime(o)?.TimeOfDay;
    }

    public static readonly Dictionary<Type, Func<object?, object?>> Converters = new()
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
}