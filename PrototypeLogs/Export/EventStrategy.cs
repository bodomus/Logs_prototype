using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PrototypeLogs.Export
{
    public class EventStrategy : BaseStrategy, IExportExcelStrategy
    {
        
        private Dictionary<int, string> _sheetHeader
        {
            get
            {
                return new Dictionary<int, string>() {
                    {1, "Date&Time"},
                    {2, "Action"},
                    {3, "Value"},
                    {4, "Description"}
                    };
            }
        }
        public EventStrategy(string excelFileName, string logFileName, uint strategyIndex) : base(excelFileName, logFileName, strategyIndex)
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
            var excel = new LogsOpenXML(_excelFileName, sheetName, _strategyIndex, true, null, false);
            excel.AddHeader(_sheetHeader);
            foreach (var s in strings)
            {
                rowIdx++;
                Row row = excel.SheetData.AppendChild(new Row() { RowIndex = rowIdx });
                var item = GetLogItem(s);
                if (item != null) {
                    excel.InsertCell(row, "", CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                    excel.InsertCell(row, item.Action, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                    excel.InsertCell(row, item.Value, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                    excel.InsertCell(row, item.Description, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                } else 
                    excel.InsertCell(row, s, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
            }
            excel.Close();
        }

        private LogItem GetLogItem(string input)
        {
            //string input = "17:09:08:148 1.0.0.0 [1] (INFO): A:MousePress V: TextBlock D: Down";

            Regex expression = new Regex(@"A:(?<Action>.*)V:(?<Value>.*)D:(?<Description>.*)$");

            Match match = expression.Match(input);

            if (match.Success)
            {
                return new LogItem()
                {
                    Action = match.Groups["Action"].Value,
                    Value = match.Groups["Value"].Value,
                    Description = match.Groups["Description"].Value
                };
            }
            return null;
        }
    }
}
