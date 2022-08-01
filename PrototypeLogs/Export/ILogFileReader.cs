using System.Collections.Generic;

namespace Pathway.WPF.ImportExport
{
    public interface ILogFileReader
    {
        IEnumerable<string> GetStrings();
    }
}
