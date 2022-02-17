using System.Collections.Generic;

namespace GoogleSheetsManager.Tests;

internal sealed class TestInstance : ISavable
{
    public IList<string> Titles => new List<string> { Title };

    public string? Value;

    public static TestInstance Load(IDictionary<string, object?> valueSet)
    {
        return new TestInstance { Value = valueSet[Title]?.ToString() };
    }

    public IDictionary<string, object?> Save() => new Dictionary<string, object?> { { Title, Value } };

    private const string Title = "Title";
}
