using System;
using System.Collections.Generic;
using JetBrains.Annotations;

namespace GoogleSheetsManager;

[PublicAPI]
public class SheetData<T>
{
    public readonly IList<T> Instances;
    public readonly IList<string> Titles;

    public SheetData() : this(Array.Empty<T>(), Array.Empty<string>()) { }
    public SheetData(IList<T> instances, IList<string> titles)
    {
        Instances = instances;
        Titles = titles;
    }
}