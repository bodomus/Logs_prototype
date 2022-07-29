using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrototypeLogs.Export
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
    }
}
