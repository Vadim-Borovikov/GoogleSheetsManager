using System.Collections.Generic;

namespace GoogleSheetsManager
{
    public interface ILoadable
    {
        void Load(IList<object> values);
    }
}