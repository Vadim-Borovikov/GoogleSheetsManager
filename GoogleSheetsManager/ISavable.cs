using System.Collections.Generic;

namespace GoogleSheetsManager;

public interface ISavable : IConvertibleTo<IDictionary<string, object?>>
{
    IList<string> Titles { get; }
}
