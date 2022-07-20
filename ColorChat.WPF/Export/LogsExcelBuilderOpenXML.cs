using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NLog;
using Pathway.WPF.ImportExport.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ColorChat.WPF.Export
{
    public sealed class LogsExcelBuilderOpenXML : ExcelBuilderOpenXML
    {
        private static Logger logger = LogManager.GetLogger("file");
        private uint rowIdx = 0;
        private int columnStepIndex = 1;
        private string LOG_EXCEPTION = "EXCEPTION";
        private string PID_EXCEPTION = "PID_EXCEPTION";
        private string EVENT_EXCEPTION = "EVENT";

        protected WorksheetPart pidSheet;

        protected SheetData pidSheetData;

        public string LOG_EXCEPTION_FILE => Path.Combine(LOG_EXCEPTION, ".xlsx");
        public string PID_EXCEPTION_FILE => Path.Combine(PID_EXCEPTION, ".xlsx");
        public string EVENT_EXCEPTION_FILE => Path.Combine(EVENT_EXCEPTION, ".xlsx");

        /// <summary>
        /// pid file create before this operation
        /// </summary>
        /// <param name="filename">represent Excel file for save data</param>
        /// <param name="logFilesName">list of log files</param>
        public LogsExcelBuilderOpenXML(string filename, List<string> logFilesName)
            : base(filename)
        {
            KeyValuePair<WorksheetPart, SheetData> dataSheets = CreateSheet("Log", 1);
            this.dataSheet = dataSheets.Key;
            this.dataSheetData = dataSheets.Value;

            KeyValuePair<WorksheetPart, SheetData> descSheets = CreateSheet("Event", 2, 25);
            this.descriptionSheet = descSheets.Key;
            this.descriptionSheetData = descSheets.Value;

            KeyValuePair<WorksheetPart, SheetData> pidSheets = CreateSheet("PID", 3, 25);
            this.pidSheet = pidSheets.Key;
            this.pidSheetData = pidSheets.Value;
        }

        /// <summary>
        /// Read all strings from log file
        /// </summary>
        /// <param name="path">path to the log file</param>
        /// <returns></returns>
        private IEnumerable<string> ReadFile(string path)
        {
            string[] readText = File.ReadAllLines(path);
            return new List<string>(readText);
        }

        /// <summary>
        /// Add Header 
        /// </summary>
        public void AddHeader()
        {
            this.rowIdx++;
            Row row = dataSheetData.AppendChild(new Row() { RowIndex = rowIdx });
                
            InsertCell(row, "Number", CellValues.String, BOLDINDEXSTYLE);
            InsertCell(row, "Value", CellValues.String, BOLDINDEXSTYLE);
            
        }

        public void DoWorkWithFile(string filePath) {
            var strings = ReadFile(filePath);

        }
    }
}
