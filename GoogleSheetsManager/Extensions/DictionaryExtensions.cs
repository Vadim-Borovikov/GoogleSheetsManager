using System.Collections.Generic;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Extensions;

[PublicAPI]
public static class DictionaryExtensions
{
    public static TValue? GetValueOrDefault<TKey, TValue>(this IDictionary<TKey, TValue> set, TKey key)
        where TKey : notnull
    {
        set.TryGetValue(key, out TValue? o);
        return o;
    }
}