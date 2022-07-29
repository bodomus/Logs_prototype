using ColorChat.WPF.EventLogger;
using DocumentFormat.OpenXml.Spreadsheet;
using PrototypeLogs.Domain;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PrototypeLogs.Export
{
    public class ExceptionStrategy : BaseStrategy, IExportExcelStrategy
    {
        private Dictionary<int, string> _sheetHeader
        {
            get
            {
                return new Dictionary<int, string>() {
                    {1, "Message"},
                    };
            }
        }
        private List<ColumnsPreference> _colunmPreferences
        {
            get
            {
                return new List<ColumnsPreference>() {
                    new ColumnsPreference{
                        Min = 1, Max = 1, Width = 200
                    }
                    };
            }
        }
        public ExceptionStrategy(string excelFileName, string logFileName, uint strategyIndex) : base(excelFileName, logFileName, strategyIndex)
        {
        }

        private string GetSheetName()
        {
            return Path.GetFileNameWithoutExtension(_logFileName);
        }
        public void DoWork()
        {
            rowIdx = 2;
            var strings = ReadFile(new LogFileTextReader(_logFileName));
            var sheetName = GetSheetName();
            var excel = new LogsOpenXML(_excelFileName, sheetName, _strategyIndex, _colunmPreferences);
            excel.AddHeader(_sheetHeader);
            
            foreach (var s in strings)
            {
                rowIdx++;
                Row row = excel.SheetData.AppendChild(new Row() { RowIndex = rowIdx });
                excel.InsertCell(row, s, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
            }
            excel.Close();
        }
    }
}
