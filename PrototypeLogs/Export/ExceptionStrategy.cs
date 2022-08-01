using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Spreadsheet;
using Pathway.WPF.ImportExport.Logs.Domain;

namespace Pathway.WPF.ImportExport.Logs.Strategies
{
    public class ExceptionStrategy : BaseStrategy, IExportExcelStrategy
    {
        private Dictionary<int, string> _sheetHeader
        {
            get
            {
                return new Dictionary<int, string>() {
                    {1, "Exception message"},
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