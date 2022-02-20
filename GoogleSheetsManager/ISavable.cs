using System.Collections.Generic;

namespace GoogleSheetsManager;

public interface ISavable : IConvertableTo<IDictionary<string, object?>>
{
    IList<string> Titles { get; }
}
