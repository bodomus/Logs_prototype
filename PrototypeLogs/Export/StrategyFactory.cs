using System;
using System.IO;

using NLog;

namespace Pathway.WPF.ImportExport.Logs.Strategies
{
    public class StrategyFactory
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
                    return new PidStrategy(excelFile, name, strategyIndex);
                }
                case "event":
                {
                    return new EventStrategy(excelFile, name, strategyIndex);
                }
                default:
                {
                    logger.Error("StrategyFactory: invald file name");
                    throw new InvalidProgramException("StrategyFactory: invald file name");
                }
            }
        }
    }
}