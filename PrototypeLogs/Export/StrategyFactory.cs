using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;

namespace PrototypeLogs.Export
{
    internal class StrategyFactory
    {
        private static Logger logger = LogManager.GetLogger("file");
        public static IExportExcelStrategy Create(string name)
        {
            switch (name)
            {
                case "LOG_EXCEPTION":
                    {
                        return new ExceptionStrategy();
                        break;
                    }

                case "LOG_PID":
                    {
                        return new ExceptionStrategy();
                        break;
                    }
                case "LOG_EVENT":
                    {
                        return new ExceptionStrategy();
                        break;
                    }
                default:
                    {
                        logger.Error("StrategyFactory invald file name");
                        throw new InvalidProgramException("StrategyFactory invald file name");
                    }
            }
            return null;
        }
    }
}
