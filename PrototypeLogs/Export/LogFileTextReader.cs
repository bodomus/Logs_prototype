﻿using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pathway.WPF.ImportExport.Logs
{
    public class LogFileTextReader : ILogFileReader
    {
        private readonly string _fileName;
        public LogFileTextReader(string filePath)
        {
            _fileName = filePath;
        } 

        public IEnumerable<string> GetStrings()
        {
            string[] readText = File.ReadAllLines(_fileName);
            return new List<string>(readText);
        }
    }
}