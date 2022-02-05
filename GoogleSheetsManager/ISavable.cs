using System.Collections.Generic;

namespace GoogleSheetsManager;

public interface ISavable
{
    IList<string> Titles { get; }
    IDictionary<string, object?> Save();
}