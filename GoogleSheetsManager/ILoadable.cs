using System.Collections.Generic;

namespace GoogleSheetsManager;

public interface ILoadable
{
    void Load(IDictionary<string, object> valueSet);
}