using JetBrains.Annotations;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;
using GoogleSheetsManager.Providers;
using System.Linq;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using Google.Apis.Sheets.v4.Data;
using GryphonUtilities.Extensions;

namespace GoogleSheetsManager.Documents;

[PublicAPI]
public class Sheet
{
    public string Name { get; private set; }

    public List<string> Titles { get; private set; } = new();

    public int? Index => _sheet?.Properties.Index;

    internal Sheet(Google.Apis.Sheets.v4.Data.Sheet sheet, SheetsProvider provider, Document document,
        IDictionary<Type, Func<object?, object?>> converters)
    : this(sheet.Properties.Title, provider, document, converters)
    {
        _sheet = sheet;
    }

    internal Sheet(string name, SheetsProvider provider, Document document,
        IDictionary<Type, Func<object?, object?>> converters)
    {
        Name = name;
        _provider = provider;
        _document = document;
        _converters = converters;
    }

    public async Task<List<T>> LoadAsync<T>(string range, bool formula = false,
        IDictionary<string, string>? titleAliases = null,
        ICollection<Func<IDictionary<string, object?>, T?, T?>>? additionalLoaders = null)
        where T : class, new()
    {
        List<Dictionary<string, object?>> maps = await LoadAsync(AddTitleTo(range), formula);
        return maps.Select(m => Load(m, titleAliases, additionalLoaders))
                   .RemoveNulls()
                   .ToList();
    }

    public async Task LoadTitlesAsync(string range)
    {
        IList<IList<object>> rows = await _provider.GetValueListAsync(AddTitleToFirstRaw(range), false);
        Titles = GetTitles(rows[0]);
    }

    public Task SaveAsync<T>(string range, List<T> instances, IDictionary<string, string>? titleAliases = null,
        IEnumerable<Action<T, IDictionary<string, object?>>>? additionalSavers = null)
    {
        List<Dictionary<string, object?>> maps =
            instances.Select(i => Save(i, titleAliases, additionalSavers)).ToList();
        return SaveAsync(AddTitleTo(range), maps);
    }

    public Task AddAsync<T>(string range, List<T> instances, IDictionary<string, string>? titleAliases = null,
        IEnumerable<Action<T, IDictionary<string, object?>>>? additionalSavers = null)
    {
        List<Dictionary<string, object?>> maps =
            instances.Select(i => Save(i, titleAliases, additionalSavers)).ToList();
        return AddAsync(AddTitleTo(range), maps);
    }

    public Task SaveRawAsync(string range, IList<IList<object>> rows)
    {
        range = AddTitleTo(range);
        return _provider.UpdateValuesAsync(range, rows);
    }

    public Task ClearAsync(string range) => _provider.ClearValuesAsync(AddTitleTo(range));

    public async Task RenameAsync(string newName)
    {
        if (_sheet is null)
        {
            Spreadsheet spreadsheet = await _document.GetSpreadsheetAsync();
            _sheet = spreadsheet.Sheets.SingleOrDefault(s => s.Properties.Title == Name);
            if (_sheet is null)
            {
                throw new NullReferenceException(nameof(_sheet));
            }
        }
        await _provider.RenameSheetAsync(_sheet.Properties.SheetId, newName);
        Name = newName;
    }

    internal void SetSheet(Google.Apis.Sheets.v4.Data.Sheet sheet) => _sheet = sheet;

    private string AddTitleTo(string range)
    {
        Range.Range fullRange = ParseAndAddName(range);
        return fullRange.ToString();
    }

    private string AddTitleToFirstRaw(string range)
    {
        Range.Range fullRange = ParseAndAddName(range);
        return fullRange.GetFirstRow().ToString();
    }

    private Range.Range ParseAndAddName(string range)
    {
        Range.Range parsed = Range.Range.Parse(range);
        Range.Range fullRange = new(parsed.IntervalStart, parsed.IntervalEnd, Name);
        return fullRange;
    }

    private async Task<List<Dictionary<string, object?>>> LoadAsync(string range, bool formula = false)
    {
        IList<IList<object>> rows = await _provider.GetValueListAsync(range, formula);
        if (rows.Count < 1)
        {
            return new List<Dictionary<string, object?>>();
        }

        Titles = GetTitles(rows[0]);
        return rows.Skip(1).Select(Organize).ToList();
    }

    private static List<string> GetTitles(IEnumerable<object> row) => row.Select(o => o.ToString() ?? "").ToList();

    private Dictionary<string, object?> Organize(IList<object> row)
    {
        Dictionary<string, object?> map = new();
        for (int i = 0; i < Titles.Count; ++i)
        {
            map[Titles[i]] = i < row.Count ? row[i] : null;
        }
        return map;
    }

    private T? Load<T>(IDictionary<string, object?> map, IDictionary<string, string>? titleAliases = null,
        IEnumerable<Func<IDictionary<string, object?>, T?, T?>>? additionalLoaders = null)
        where T : class, new()
    {
        T? instance = new();
        Type type = typeof(T);
        instance = Load(map, titleAliases, instance, type.GetProperties(), info => info.PropertyType,
            (info, obj, val) => info.SetValue(obj, val));
        instance = Load(map, titleAliases, instance, type.GetFields(), info => info.FieldType,
            (info, obj, val) => info.SetValue(obj, val));
        if (additionalLoaders is not null)
        {
            // ReSharper disable once LoopCanBeConvertedToQuery
            foreach (Func<IDictionary<string, object?>, T?, T?> loader in additionalLoaders)
            {
                instance = loader(map, instance);
            }
        }
        return instance;
    }

    private TInstance? Load<TInstance, TInfo>(IDictionary<string, object?> map,
        IDictionary<string, string>? titleAliases, TInstance? instance, IEnumerable<TInfo> members,
        Func<TInfo, Type> typeProvider, Action<TInfo, TInstance, object?> setter)
        where TInstance : class
        where TInfo : MemberInfo
    {
        if (instance is null)
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
                        title = GetTitle(sheetField, titleAliases, info);
                        break;
                }
            }

            if (title is null)
            {
                continue;
            }

            Type type = typeProvider(info);
            Func<object?, object?>? converter = _converters.AsReadOnly().GetValueOrDefault(type);
            object? converted = map.TryGetValue(title, out object? value) ? converter?.Invoke(value) : null;
            if (required && converted is null or "")
            {
                return null;
            }
            setter(info, instance, converted);
        }

        return instance;
    }

    private Task SaveAsync(string range, List<Dictionary<string, object?>> maps)
    {
        List<IList<object>> rows = new() { Titles.ToList<object>() };
        rows.AddRange(maps.Select(set => Titles.Select(t => set.GetValueOrDefault(t) ?? "").ToList()));
        return _provider.UpdateValuesAsync(range, rows);
    }

    private Task AddAsync(string range, List<Dictionary<string, object?>> maps)
    {
        List<IList<object>> rows =
            new(maps.Select(set => Titles.Select(t => set.GetValueOrDefault(t) ?? "").ToList()));
        return _provider.AppendValuesAsync(range, rows);
    }

    private static Dictionary<string, object?> Save<T>(T instance, IDictionary<string, string>? titleAliases = null,
        IEnumerable<Action<T, IDictionary<string, object?>>>? additionalSavers = null)
    {
        Dictionary<string, object?> map = new();
        Type type = typeof(T);
        Save(instance, map, type.GetProperties(), (info, obj) => info.GetValue(obj), titleAliases);
        Save(instance, map, type.GetFields(), (info, obj) => info.GetValue(obj), titleAliases);
        if (additionalSavers is not null)
        {
            foreach (Action<T, IDictionary<string, object?>> saver in additionalSavers)
            {
                saver(instance, map);
            }
        }
        return map;
    }

    private static void Save<TInstance, TInfo>(TInstance instance, Dictionary<string, object?> map,
        IEnumerable<TInfo> members, Func<TInfo, TInstance, object?> getter,
        IDictionary<string, string>? titleAliases = null)
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
            string title = GetTitle(sheetField, titleAliases, info);
            map[title] = sheetField.Format is null ? value : string.Format(sheetField.Format, value);
        }
    }

    private static string GetTitle(SheetFieldAttribute sheetField, IDictionary<string, string>? titleAliases,
        MemberInfo info)
    {
        if (sheetField.Title is not null)
        {
            return sheetField.Title;
        }

        if (titleAliases is not null && titleAliases.ContainsKey(info.Name))
        {
            return titleAliases[info.Name];
        }

        return info.Name;
    }

    private readonly SheetsProvider _provider;
    private readonly Document _document;
    private readonly IDictionary<Type, Func<object?, object?>> _converters;

    private Google.Apis.Sheets.v4.Data.Sheet? _sheet;
}