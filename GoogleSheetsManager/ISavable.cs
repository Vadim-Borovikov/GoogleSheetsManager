using System.Collections.Generic;

namespace GoogleSheetsManager
{
    public interface ISavable
    {
        IList<object> Save();
    }
}