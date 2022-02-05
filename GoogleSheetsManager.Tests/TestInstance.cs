using System.Collections.Generic;

namespace GoogleSheetsManager.Tests;

internal sealed class TestInstance : ILoadable, ISavable
{
    public IList<string> Titles => new List<string> { Title };

    public string? Value;

    public void Load(IDictionary<string, object?> valueSet) => Value = valueSet[Title]?.ToString();

    public IDictionary<string, object?> Save() => new Dictionary<string, object?> { { Title, Value } };

    private const string Title = "Title";
}