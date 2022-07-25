using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrototypeLogs.Export
{
    public class ExceptionStrategy :BaseStrategy, IExportExcelStrategy
    {
        private string _logFileName;
        public ExceptionStrategy(SheetData data, string logFileName) : base(data) {
            _logFileName = logFileName;
        }
        private uint rowIdx = 0;
        public void DoWork(
            )
        {
            rowIdx = 2;
            var strings = ReadFile(logFile);
            foreach (var s in strings)
            {

                rowIdx++;
                Row row = Data.AppendChild(new Row() { RowIndex = rowIdx });
                //InsertCell(row, s, CellValues.String, BOLDINDEXSTYLE);
            }
        }
    }
}
