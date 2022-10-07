using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using GryphonUtilities;
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

            result[sheetField.Title] = getter(info, instance);
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
            object? value = Convert(valueSet[title], type);
            if (required && value is null or "")
            {
                return null;
            }
            setter(info, result, value);
        }

        return result;
    }

    private static object? Convert(object? value, Type type)
    {
        if ((type == typeof(bool)) || (type == typeof(bool?)))
        {
            return value.ToBool();
        }
        if ((type == typeof(byte)) || (type == typeof(byte?)))
        {
            return value.ToByte();
        }
        if ((type == typeof(ushort)) || (type == typeof(ushort?)))
        {
            return value.ToUshort();
        }
        if ((type == typeof(int)) || (type == typeof(int?)))
        {
            return value.ToInt();
        }
        if ((type == typeof(decimal)) || (type == typeof(decimal?)))
        {
            return value.ToDecimal();
        }
        if ((type == typeof(long)) || (type == typeof(long?)))
        {
            return value.ToLong();
        }
        if (type == typeof(Uri))
        {
            return value.ToUri();
        }
        if (type == typeof(List<Uri>))
        {
            return value.ToUris();
        }
        if ((type == typeof(DateTime)) || (type == typeof(DateTime?)))
        {
            return value.ToDateTime();
        }
        if ((type == typeof(TimeSpan)) || (type == typeof(TimeSpan?)))
        {
            return value.ToTimeSpan();
        }

        return value;
    }

    private static bool? ToBool(this object? o)
    {
        if (o is bool b)
        {
            return b;
        }
        return bool.TryParse(o?.ToString(), out b) ? b : null;
    }

    private static byte? ToByte(this object? o)
    {
        if (o is byte b)
        {
            return b;
        }
        return byte.TryParse(o?.ToString(), out b) ? b : null;
    }

    private static ushort? ToUshort(this object? o)
    {
        if (o is ushort u)
        {
            return u;
        }
        return ushort.TryParse(o?.ToString(), out u) ? u : null;
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

    private static long? ToLong(this object? o)
    {
        if (o is long l)
        {
            return l;
        }
        return long.TryParse(o?.ToString(), out l) ? l : null;
    }

    private static Uri? ToUri(this object? o)
    {
        if (o is Uri uri)
        {
            return uri;
        }
        string? uriString = o?.ToString();
        return string.IsNullOrWhiteSpace(uriString) ? null : new Uri(uriString);
    }

    private static List<Uri>? ToUris(this object? o)
    {
        if (o is IEnumerable<Uri> l)
        {
            return l.ToList();
        }
        return o?.ToString()?.Split("\n").Select(ToUri).RemoveNulls().ToList();
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
}