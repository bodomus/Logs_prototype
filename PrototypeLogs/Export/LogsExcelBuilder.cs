using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrototypeLogs.Export
{
    public class LogsExcelBuilder
    {
        public string fileName { get; set; } 
        public string sheetName { get; set; } 
        public uint sheetIndex { get; set; } 
        public bool isReadonly { get; set; }
}
}
