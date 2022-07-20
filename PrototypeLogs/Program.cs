using ColorChat.WPF.Export;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

            logger.Error("Error");
            logger.Info("Info");
            logger.Trace("Info");
            logger1.Error("Error1");
            logger1.Info("Info1");  
            logger1.Trace("Info1");
            var logsFiles = LogsExporter.GetLogs();
            var excelFile = LogsExporter.GetExcelFileName();
            var logExport = new LogsExporter(logsFiles, excelFile);
        }
    }
}
