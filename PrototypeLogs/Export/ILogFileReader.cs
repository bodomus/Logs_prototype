using System.Collections.Generic;

namespace PrototypeLogs.Export
{
    public interface ILogFileReader
    {
        IEnumerable<string> GetStrings();
    }
}
