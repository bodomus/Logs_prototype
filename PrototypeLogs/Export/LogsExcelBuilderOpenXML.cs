using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PrototypeLogs.Export;
using NLog;
using Pathway.WPF.ImportExport;

namespace ColorChat.WPF.Export
{
    public sealed class LogsExcelBuilderOpenXML
    {
        private static Logger logger = LogManager.GetLogger("file");
        private IExportExcelStrategy _currentStrategy;
        private int _currentIndexStrategy;
        
        private string LOG_EXCEPTION = "LOG_EXCEPTION";
        private string PID_EXCEPTION = "LOG_PID";
        private string EVENT_EXCEPTION = "LOG_EVENT";
        private List<string> _logFilesNames;
        private string _excelFile;
        protected IExportExcelStrategy exportExcelStrategy;
        public event ProgressEventDelegate ProgressExportLog;
        /// <summary>
        /// pid file create before this operation
        /// </summary>
        /// <param name="filename">represent Excel file for save data</param>
        /// <param name="logFilesName">list of log files</param>
        public LogsExcelBuilderOpenXML(string filename, List<string> logFilesName)//: base(filename)
        {
            _excelFile = filename;
            _currentIndexStrategy = 0;
            _logFilesNames = logFilesName;
        }

        public int GetCurrentStrategy()
        {
            return _currentIndexStrategy;
        }

        public void SetCurrentStrategy(int v)
        {
            string fileName = _logFilesNames[v - 1];
            IExportExcelStrategy st = StrategyFactory.Create(_excelFile, fileName, (uint)v);
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
            if (stindex < 0)
                return;
            do
            {
                SetCurrentStrategy(stindex);
                _currentStrategy.DoWork();

                OnProgress((double)stindex / (double)_logFilesNames.Count);

                stindex = GetNextStrategy();
            } while (stindex > 0); 
        }

        private void OnProgress(double progress) {
            if (ProgressExportLog != null)
            {
                ProgressExportLog(this, progress);
            }
        }
    }
}
