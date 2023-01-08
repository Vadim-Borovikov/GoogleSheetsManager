using JetBrains.Annotations;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;
using GoogleSheetsManager.Providers;
using System.Linq;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using GoogleSheetsManager.Extensions;
using Google.Apis.Sheets.v4.Data;
using GryphonUtilities.Extensions;

namespace GoogleSheetsManager.Documents;

[PublicAPI]
public class Sheet
{
    public string Title { get; private set; }

    public int? Index => _sheet?.Properties.Index;

    internal Sheet(Google.Apis.Sheets.v4.Data.Sheet sheet, SheetsProvider provider, Document document,
        IDictionary<Type, Func<object?, object?>> converters)
    : this(sheet.Properties.Title, provider, document, converters)
    {
        _sheet = sheet;
    }

    internal Sheet(string title, SheetsProvider provider, Document document,
        IDictionary<Type, Func<object?, object?>> converters)
    {
        Title = title;
        _provider = provider;
        _document = document;
        _converters = converters;
    }

    public async Task<SheetData<T>> LoadAsync<T>(string range, bool formula = false,
        ICollection<Func<IDictionary<string, object?>, T?, T?>>? additionalLoaders = null)
        where T : class, new()
    {
        SheetData<Dictionary<string, object?>> data = await LoadAsync(AddTitleTo(range), formula);
        return Load(data, additionalLoaders);
    }

    public async Task<List<string>> LoadTitlesAsync(string range)
    {
        IList<IList<object>> rawValueSets = await _provider.GetValueListAsync(AddTitleToFirstRaw(range), false);
        return GetTitles(rawValueSets[0]).ToList();
    }

    public Task SaveAsync<T>(string range, SheetData<T> data,
        IEnumerable<Action<T, IDictionary<string, object?>>>? additionalSavers = null)
    {
        List<Dictionary<string, object?>> instances = data.Instances.Select(i => Save(i, additionalSavers)).ToList();
        SheetData<Dictionary<string, object?>> savedData = new(instances, data.Titles);
        return SaveAsync(AddTitleTo(range), savedData);
    }

    public Task ClearAsync(string range) => _provider.ClearValuesAsync(AddTitleTo(range));

    public async Task RenameAsync(string newName)
    {
        if (_sheet is null)
        {
            Spreadsheet spreadsheet = await _document.GetSpreadsheetAsync();
            _sheet = spreadsheet.Sheets.SingleOrDefault(s => s.Properties.Title == Title);
            if (_sheet is null)
            {
                throw new NullReferenceException(nameof(_sheet));
            }
        }
        await _provider.RenameSheetAsync(_sheet.Properties.SheetId, newName);
        Title = newName;
    }

    internal void SetSheet(Google.Apis.Sheets.v4.Data.Sheet sheet) => _sheet = sheet;

    private string AddTitleTo(string range)
    {
        Range.Range fullRange = ParseAndAddTitle(range);
        return fullRange.ToString();
    }

    private string AddTitleToFirstRaw(string range)
    {
        Range.Range fullRange = ParseAndAddTitle(range);
        return fullRange.GetFirstRow().ToString();
    }

    private Range.Range ParseAndAddTitle(string range)
    {
        Range.Range parsed = Range.Range.Parse(range);
        Range.Range fullRange = new(parsed.IntervalStart, parsed.IntervalEnd, Title);
        return fullRange;
    }

    private async Task<SheetData<Dictionary<string, object?>>> LoadAsync(string range, bool formula = false)
    {
        IList<IList<object>> rawValueSets = await _provider.GetValueListAsync(range, formula);
        if (rawValueSets.Count < 1)
        {
            return new SheetData<Dictionary<string, object?>>();
        }
        List<string> titles = GetTitles(rawValueSets[0]).ToList();

        List<Dictionary<string, object?>> valueSets = rawValueSets.Skip(1).Select(r => Organize(r, titles)).ToList();
        return new SheetData<Dictionary<string, object?>>(valueSets, titles);
    }

    private static IEnumerable<string> GetTitles(IEnumerable<object> rawValueSet)
    {
        return rawValueSet.Select(o => o.ToString() ?? "");
    }

    private static Dictionary<string, object?> Organize(IList<object> rawValueSet, IList<string> titles)
    {
        Dictionary<string, object?> result = new();
        for (int i = 0; i < titles.Count; ++i)
        {
            result[titles[i]] = i < rawValueSet.Count ? rawValueSet[i] : null;
        }
        return result;
    }

    private SheetData<T> Load<T>(SheetData<Dictionary<string, object?>> data,
        ICollection<Func<IDictionary<string, object?>, T?, T?>>? additionalLoaders = null)
        where T : class, new()
    {
        List<T> instances = data.Instances
                                .Select(set => Load(set, additionalLoaders))
                                .RemoveNulls()
                                .ToList();
        return new SheetData<T>(instances, data.Titles);
    }

    private T? Load<T>(IDictionary<string, object?> valueSet,
        IEnumerable<Func<IDictionary<string, object?>, T?, T?>>? additionalLoaders = null)
        where T : class, new()
    {
        T? result = new();
        Type type = typeof(T);
        result = Load(valueSet, result, type.GetProperties(), info => info.PropertyType,
            (info, obj, val) => info.SetValue(obj, val));
        result = Load(valueSet, result, type.GetFields(), info => info.FieldType,
            (info, obj, val) => info.SetValue(obj, val));
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

    private TInstance? Load<TInstance, TInfo>(IDictionary<string, object?> valueSet, TInstance? result,
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
            Func<object?, object?>? converter = _converters.GetValueOrDefault(type);
            object? value = valueSet.TryGetValue(title, out object? rawValue) ? converter?.Invoke(rawValue) : null;
            if (required && value is null or "")
            {
                return null;
            }
            setter(info, result, value);
        }

        return result;
    }

    private Task SaveAsync(string range, SheetData<Dictionary<string, object?>> data)
    {
        List<IList<object>> rawValueSets = new() { data.Titles.ToList<object>() };
        rawValueSets.AddRange(data.Instances
                                  .Select(set => data.Titles
                                                     .Select(t => DictionaryExtensions.GetValueOrDefault(set, t) ?? "")
                                                     .ToList()));
        return _provider.UpdateValuesAsync(range, rawValueSets);
    }

    private static Dictionary<string, object?> Save<T>(T instance,
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

    private static void Save<TInstance, TInfo>(TInstance instance, Dictionary<string, object?> result,
        IEnumerable<TInfo> members, Func<TInfo, TInstance, object?> getter)
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

    private readonly SheetsProvider _provider;
    private readonly Document _document;
    private readonly IDictionary<Type, Func<object?, object?>> _converters;

    private Google.Apis.Sheets.v4.Data.Sheet? _sheet;
}