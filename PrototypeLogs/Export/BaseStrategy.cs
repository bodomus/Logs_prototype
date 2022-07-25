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
        public SheetData Data;

        public BaseStrategy() {
            //Data = data;
        }
        /// <summary>
        /// Read all strings from log file
        /// </summary>
        /// <param name="path">path to the log file</param>
        /// <returns></returns>
        protected IEnumerable<string> ReadFile(string path)
        {
            string[] readText = File.ReadAllLines(path);
            return new List<string>(readText);
        }
    }
}
