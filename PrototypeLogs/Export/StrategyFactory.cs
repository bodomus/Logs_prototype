using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;

namespace PrototypeLogs.Export
{
    internal class StrategyFactory
    {
        private static Logger logger = LogManager.GetLogger("file");
        public static IExportExcelStrategy Create(string excelFile, string name, uint strategyIndex)
        {
            switch (Path.GetFileNameWithoutExtension(name))
            {
                case "short":
                    {
                        return new ExceptionStrategy(excelFile, name, strategyIndex);
                    }

                case "pid":
                    {
                        return new EventStrategy(excelFile, name, strategyIndex);
                    }
                case "event":
                    {
                        return new ExceptionStrategy(excelFile, name, strategyIndex);
                    }
                default:
                    {
                        logger.Error("StrategyFactory invald file name");
                        throw new InvalidProgramException("StrategyFactory invald file name");
                    }
            }
        }
    }
}
