using System.Collections.Generic;

using DocumentFormat.OpenXml.Spreadsheet;
using Pathway.WPF.ImportExport.Logs.Domain;
using Pathway.WPF.ImportExport.Logs.StyleSheetBuilder;

namespace Pathway.WPF.ImportExport.Logs.Strategies
{
    public class ExceptionStrategy : BaseStrategy, IExportExcelStrategy
    {
        private Dictionary<int, string> _sheetHeader
        {
            get
            {
                return new Dictionary<int, string>() {
                    {1, "№№"},
                    {2, "Messages"},
                };
            }
        }
        private List<ColumnsPreference> _colunmPreferences
        {
            get
            {
                return new List<ColumnsPreference>() {
                    new ColumnsPreference{
                        Min = 1, Max = 1, Width = 20
                    },
                    new ColumnsPreference{
                        Min = 2, Max = 2, Width = 500
                    }
                };
            }
        }
        public ExceptionStrategy(string excelFileName, string logFileName, uint strategyIndex) : base(excelFileName, logFileName, strategyIndex)
        {
        }

        public void DoWork()
        {
            rowIdx = 1;
            var strings = ReadFile(new LogFileTextReader(_logFileName));
            var sheetName = GetSheetName();
            var excel = new LogsOpenXML(_excelFileName, sheetName, _strategyIndex, _colunmPreferences);
            
            IStyleSheetWorker worker = new StyleSheetExceptionWorker(); 
            worker.Prepare(excel.Stylesheet);
            excel.Worker = worker;
            excel.AddHeader(_sheetHeader, worker.IndexRefCellHeaderBase);

            // excel.AddHeader(_sheetHeader);
            foreach (var s in strings)
            {
                rowIdx++;
                Row row = excel.SheetData.AppendChild(new Row() { RowIndex = rowIdx });
                excel.FormatCell(row, "A", (rowIdx - 1).ToString());
                excel.FormatCell(row, "B", s);
                // excel.InsertCell(row, s, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                // excel.InsertCell(row, s, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
            }
            excel.Close();
        }
    }
}