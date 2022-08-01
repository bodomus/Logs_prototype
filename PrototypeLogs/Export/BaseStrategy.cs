using System.Collections.Generic;
using System.IO;

namespace Pathway.WPF.ImportExport.Logs.Strategies
{
    public class BaseStrategy
    {
        protected uint rowIdx;

        
        protected string _logFileName;
        protected string _excelFileName;
        protected uint _strategyIndex;

        public BaseStrategy(string excelFileName, string logFileName, uint strategyIndex) {
            _logFileName = logFileName;
            _excelFileName = excelFileName;
            _strategyIndex = strategyIndex;
            rowIdx = 0;
        }

        /// <summary>
        /// Read all strings from log file
        /// </summary>
        /// <param name="path">path to the log file</param>
        /// <returns></returns>
        protected IEnumerable<string> ReadFile(ILogFileReader logFileReader)
        {
            return logFileReader?.GetStrings();
        }

        protected string GetSheetName()
        {
            return Path.GetFileNameWithoutExtension(_logFileName);
        }
    }
}