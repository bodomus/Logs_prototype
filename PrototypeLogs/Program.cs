using ColorChat.WPF.Export;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PrototypeLogs
{
    internal class Program
    {
        private static Logger logger = LogManager.GetLogger("file");
        private static Logger logger1 = LogManager.GetLogger("file1");
        //private EventLoggerClass EL = new EventLoggerClass();

        static void Main(string[] args)
        {
            Test();
            var logsFiles = LogsExporter.GetLogs();
            var excelFile = LogsExporter.GetExcelFileName();
            var logExport = new LogsExporter(logsFiles, excelFile);
        }

        public static void Test() {
            
            string input = "11:19:54:607 1.0.0.0 [1] (INFO): A:TI TextBox TimeStamp: 11405343 V: 2 ";

            // ... Use named group in regular expression.
            Regex expression = new Regex(@"A:(?<Action>\.*)V:");

            // ... See if we matched.
            Match match = expression.Match(input);
            if (match.Success)
            {
                // ... Get group by name.
                string result = match.Groups["Action"].Value;
                Console.WriteLine("Action: {0}", result);
            }
            // Done.
            Console.ReadLine();
        }
    }
}
