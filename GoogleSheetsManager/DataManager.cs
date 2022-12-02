using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using GoogleSheetsManager.Providers;
using System.Threading.Tasks;
using GryphonUtilities;
using JetBrains.Annotations;

namespace GoogleSheetsManager;

[PublicAPI]
public static class DataManager<T>
    where T : class, new()
{
    public static async Task<SheetData<T>> LoadAsync(SheetsProvider provider, string range, int? sheetIndex = null,
        bool formula = false, IDictionary<Type, Func<object?, object?>>? additionalConverters = null,
        ICollection<Func<IDictionary<string, object?>, T?, T?>>? additionalLoaders = null)
    {
        SheetData<Dictionary<string, object?>> data = await Utils.LoadAsync(provider, range, sheetIndex, formula);
        return Load(data, provider.TimeManager, additionalConverters, additionalLoaders);
    }

    public static Task SaveAsync(SheetsProvider provider, string range, SheetData<T> data,
        IEnumerable<Action<T, IDictionary<string, object?>>>? additionalSavers = null)
    {
        List<Dictionary<string, object?>> instances = data.Instances.Select(i => Save(i, additionalSavers)).ToList();
        SheetData<Dictionary<string, object?>> savedData = new(instances, data.Titles);
        return Utils.SaveAsync(provider, range, savedData);
    }

    private static SheetData<T> Load(SheetData<Dictionary<string, object?>> data, TimeManager timeManager,
        IDictionary<Type, Func<object?, object?>>? additionalConverters = null,
        ICollection<Func<IDictionary<string, object?>, T?, T?>>? additionalLoaders = null)
    {
        if (data.Instances.Count < 1)
        {
            return new SheetData<T>();
        }

        List<T> instances = data.Instances
                                .Select(set => Load(set, timeManager, additionalConverters, additionalLoaders))
                                .RemoveNulls()
                                .ToList();
        return new SheetData<T>(instances, data.Titles);
    }

    private static T? Load(IDictionary<string, object?> valueSet, TimeManager timeManager,
        IDictionary<Type, Func<object?, object?>>? additionalConverters = null,
        IEnumerable<Func<IDictionary<string, object?>, T?, T?>>? additionalLoaders = null)
    {
        T? result = new();
        Type type = typeof(T);
        result = Load(valueSet, result, type.GetProperties(), info => info.PropertyType,
            (info, obj, val) => info.SetValue(obj, val), additionalConverters, timeManager);
        result = Load(valueSet, result, type.GetFields(), info => info.FieldType,
            (info, obj, val) => info.SetValue(obj, val), additionalConverters, timeManager);
        if (additionalLoaders is not null)
        {
            // ReSharper disable once LoopCanBeConvertedToQuery
            foreach (Func<IDictionary<string, object?>, T?, T?> loader in additionalLoaders)
            {
                result = loader(valueSet, result);
            }
        }
        return result;
    }

    private static TInstance? Load<TInstance, TInfo>(IDictionary<string, object?> valueSet, TInstance? result,
        IEnumerable<TInfo> members, Func<TInfo, Type> typeProvider, Action<TInfo, TInstance, object?> setter,
        IDictionary<Type, Func<object?, object?>>? additionalConverters, TimeManager timeManager)
        where TInstance : class
        where TInfo : MemberInfo
    {
        if (result is null)
        {
            return null;
        }

        Dictionary<Type, Func<object?, object?>> converters = new(Utils.DefaultConverters);
        converters[typeof(DateTimeFull)] = converters[typeof(DateTimeFull?)] =
            o => Utils.GetDateTimeFull(o, timeManager);
        if (additionalConverters is not null)
        {
            foreach (Type type in additionalConverters.Keys)
            {
                converters[type] = additionalConverters[type];
            }
        }

        foreach (TInfo info in members)
        {
            bool required = false;
            string? title = null;
            foreach (Attribute attribute in info.GetCustomAttributes())
            {
                switch (attribute)
                {
                    case RequiredAttribute:
                        required = true;
                        break;
                    case SheetFieldAttribute sheetField:
                        title = sheetField.Title ?? info.Name;
                        break;
                }
            }

            if (title is null)
            {
                continue;
            }

            Type type = typeProvider(info);
            Func<object?, object?>? converter = converters.GetValueOrDefault(type);
            object? value = valueSet.TryGetValue(title, out object? rawValue) ? converter?.Invoke(rawValue) : null;
            if (required && value is null or "")
            {
                return null;
            }
            setter(info, result, value);
        }

        return result;
    }
    private static Dictionary<string, object?> Save(T instance,
        IEnumerable<Action<T, IDictionary<string, object?>>>? additionalSavers = null)
    {
        Dictionary<string, object?> result = new();
        Type type = typeof(T);
        Save(instance, result, type.GetProperties(), (info, obj) => info.GetValue(obj));
        Save(instance, result, type.GetFields(), (info, obj) => info.GetValue(obj));
        if (additionalSavers is not null)
        {
            foreach (Action<T, IDictionary<string, object?>> saver in additionalSavers)
            {
                saver(instance, result);
            }
        }
        return result;
    }

    private static void Save<TInfo>(T instance, Dictionary<string, object?> result, IEnumerable<TInfo> members,
        Func<TInfo, T, object?> getter)
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
            string title = sheetField.Title ?? info.Name;
            result[title] = sheetField.Format is null ? value : string.Format(sheetField.Format, value);
        }
    }
}