using System.Collections.Generic;
using JetBrains.Annotations;

namespace GoogleSheetsManager.Documents;

[PublicAPI]
public sealed record SheetLoadedData<T>
{
    public readonly List<T> Instances = new();
    public readonly List<string> Titles = new();

    internal SheetLoadedData() { }

    internal SheetLoadedData(List<T> instances, List<string> titles)
    {
        Instances = instances;
        Titles = titles;
    }
}