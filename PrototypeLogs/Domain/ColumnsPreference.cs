using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pathway.WPF.ImportExport.Logs.Domain
{
    public class ColumnsPreference
    {
        public uint Min { get; set; }
        public uint Max { get; set; }
        public DoubleValue Width { get; set; }

    }
}
