using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrototypeLogs.Export
{
    public class ExceptionStrategy3 : BaseStrategy, IExportExcelStrategy
    {
        private string _logFileName;
        private string _excelFileName;
        private uint _strategyIndex;
        private uint rowIdx = 0;
        private Dictionary<int, string> _sheetHeader
        {
            get
            {
                return new Dictionary<int, string>() {
                    {1, "Name"},
                    {2, "Name 2"},
                    {3, "Name 3"},
                    {4, "Name 4"}
                    };
            }
        }
        public ExceptionStrategy3(string excelFile, string logFileName, uint strategyIndex) : base()
        {
            _logFileName = logFileName;
            _excelFileName = excelFile;
            _strategyIndex = strategyIndex;
        }

        private string GetSheetName()
        {
            return Path.GetFileNameWithoutExtension(_logFileName);
        }
        public void DoWork()
        {
            rowIdx = 2;
            var strings = ReadFile(_logFileName);
            var sheetName = GetSheetName();
            var excel = new OpenXML(_excelFileName, sheetName, _strategyIndex, true, false);
            excel.SetColumnWidth(1, 10, 200);
            excel.AddHeader(_sheetHeader);
            //excel.SetCurrentSheetByName(sheetName);
            foreach (var s in strings)
            {
                rowIdx++;
                Row row = excel.SheetData.AppendChild(new Row() { RowIndex = rowIdx });
                excel.InsertCell(row, s, CellValues.String, ExcelConstant.BOLDINDEXSTYLE);
            }
            excel.Close();
        }
    }
}
