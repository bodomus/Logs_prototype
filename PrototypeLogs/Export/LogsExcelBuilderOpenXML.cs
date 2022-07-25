using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NLog;
using Pathway.WPF.ImportExport.Excel;
using PrototypeLogs.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ColorChat.WPF.Export
{
    public sealed class LogsExcelBuilderOpenXML : ExcelBuilderOpenXML
    {
        private static Logger logger = LogManager.GetLogger("file");
        private IExportExcelStrategy _currentStrategy;
        private int _currentIndexStrategy;
        
        private int columnStepIndex = 1;
        private string LOG_EXCEPTION = "LOG_EXCEPTION";
        private string PID_EXCEPTION = "LOG_PID";
        private string EVENT_EXCEPTION = "LOG_EVENT";
        private List<string> _logFilesNames;

        protected WorksheetPart pidSheet;
        protected IExportExcelStrategy exportExcelStrategy;

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
            _currentIndexStrategy = 0;
            _logFilesNames = logFilesName;
            //LogsExcelBuilderOpenXML.CreateSpreadsheetWorkbook(filename);
        }

        public int GetCurrentStrategy()
        {
            return _currentIndexStrategy;
        }
        public void SetCurrentStrategy(int v)
        {
            string fileName = _logFilesNames[v];
            IExportExcelStrategy st = StrategyFactory.Create(fileName);
            SetStrategy(st);
        }


        public int GetNextStrategy() {
            _currentIndexStrategy++;
            if (_currentIndexStrategy > _logFilesNames.Count)
                return -1;
            return _currentIndexStrategy;
        }

        void SetStrategy(IExportExcelStrategy strategy)
        {
            if (strategy == null)
                logger.Error("Could not  get current strategy.");
            _currentStrategy = strategy;
        }

        public void DoWorkWithFile(string filePath)
        {
            var stindex = GetNextStrategy();
            do {
                SetCurrentStrategy(stindex);
                _currentStrategy.DoWork();
                stindex = GetNextStrategy();
            } while () 
            
            SaveAndClose();
        }

              
    }
}
