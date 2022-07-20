using Medoc.Configuration;
using NLog;
using Pathway.WPF.ImportExport;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;

namespace ColorChat.WPF.Export
{
    public class LogsExporter : IExporter
    {
        private static string fileLog = "short.log";
        private static string pidLog = "pid.log";
        private static string eventLog = "event.log";
        private static string excelFile = "Excel.xlsx";

        private static Logger logger = LogManager.GetLogger("file");
        private static Logger logger1 = LogManager.GetLogger("file1");
        private static int _ref;
        /// <summary>
        /// Thread for export data
        /// </summary>
        private Thread m_exportThread;
        private static LogsExcelBuilderOpenXML _builder; //builder for Excel file

        protected List<string> m_logList;
        protected string m_fileName;
        protected IExporter m_currentExporter;
        protected int m_currentInnerResultIdx;

        /// <summary>
        /// Raised during export process
        /// </summary>
        public event ProgressEventDelegate Progress;

        /// <summary>
        /// Raised when export aborted
        /// </summary>
        public event ErrorEventDelegate Aborted;

        /// <summary>
        /// Raised when export finished
        /// </summary>
        public event BasicDelegate Finished;

        /// <summary>
        /// Raised during export process
        /// </summary>
        public event MessageEventDelegate ProgressMessage;


        private object m_locker = new object();

        private static string GetDateTimeDirectory() { 
            return DateTime.Now.ToString("yyyy-MM-dd");
        }

        public static List<string> GetLogs() {
            var prop = logger.Properties;
            var path = new FileInfo(Assembly.GetEntryAssembly().Location).Directory.FullName;
            
            //($"{path}/Logs/Medoc.Remote.SignalR-{{Date}}.txt");
            var dt = GetDateTimeDirectory();
            var directory = Path.Combine(path, "Logs", dt);
            List<string> files = new List<string> {
                Path.Combine(directory, fileLog),
                Path.Combine(directory, pidLog),
                Path.Combine(directory, eventLog),
            };

            return files;
        }

        public static string GetExcelFileName()
        {
            var prop = logger.Properties;
            var path = new FileInfo(Assembly.GetEntryAssembly().Location).Directory.FullName;
            var file = Path.Combine(path, "Logs", excelFile);
            //($"{path}/Logs/Medoc.Remote.SignalR-{{Date}}.txt");
            return file;
        }

        /// <summary>
        /// Creates BatchTestResultsExporter
        /// </summary>
        /// <param name="logList">List of results</param>
        /// <param name="fileName">Name of output file of export</param>
        public LogsExporter(List<string> logList, string fileName)
        {
            if (logList == null)
                throw new ArgumentNullException("logList");

            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentException("File name can't be empty", "fileName");

            this.m_logList = logList;
            this.m_fileName = fileName;
            if (File.Exists(this.m_fileName))
            {
                try
                {
                    File.Delete(this.m_fileName);
                }
                catch (IOException)
                {
                    this.OnExportAborted(new Exception("The destination file is being used"));
                    return;
                }
            }
            if (_builder == null)
            {
                _builder = new LogsExcelBuilderOpenXML(this.m_fileName, m_logList);
                logger1.Info($"_builder is created.");
            }
            Interlocked.Increment(ref _ref);
            StartProcess();
        }

        private void StartProcess()
        {
            //TODO REmove test line
            this.ExportThread();
            logger1.Info($"StartProcess is running.");
            lock (this.m_locker)
            {
                if (this.m_exportThread != null && this.m_exportThread.IsAlive)
                    throw new InvalidOperationException("Export is already started");

                this.m_exportThread = new Thread(new ThreadStart(this.ExportThread));
                this.m_exportThread.Name = "ThermodeTestResultsExcelExporter";
                this.m_exportThread.IsBackground = true;
                this.m_exportThread.Start();
            }
        }


        private void ExportThread()
        {
            logger1.Info($"ExportThread is running.");
            

            try
            {
                //this.ReportProgress(ResourcesServices.GetString("WritingToFileMsg", this.m_fileName));

                LogsExcelBuilderOpenXML builder = new LogsExcelBuilderOpenXML(this.m_fileName, m_logList);
                lock (this.m_locker)
                {
                    if (_builder != null)
                        throw new InvalidOperationException("Export is already started");

                    //this.FreeCurrentExporter();


                    this.RegisterExporterEvents();
                    this.Start();

                    this.OnProgress(1);
                    this.OnFinish();

                    return;

                }



                //this.ExportCovasData(this.m_results, builder);
                //this.ExportStatisticsData(this.m_results, builder);

                builder.SaveAndClose();

                this.OnFinished();
            }
            catch (ThreadAbortException)
            { }
            catch (Exception e)
            {
                logger.ErrorException(e.Message, e);
                this.OnExportAborted(e);
            }
        }

        /// <summary>
        /// Start export
        /// </summary>
        public void Start()
        {

        }

        /// <summary>
        /// Starts processing of the next inner result
        /// </summary>
        /// <returns>
        /// True if next result processing started, false if nothing more to do
        /// </returns>
        //protected virtual bool ProcessNextInnerResult()
        //{
        //	lock (this.m_locker)
        //	{
        //		this.FreeCurrentExporter();

        //		if (++this.m_currentInnerResultIdx >= this.m_logList.Count)
        //			return false;

        //		this.m_currentExporter = ResultsExcelExporterFactory.Create(result, description, this.GetFileName(result));

        //		this.RegisterExporterEvents();
        //		this.m_currentExporter.Start();
        //	}

        //	return true;
        //}


        private void RegisterExporterEvents()
        {
            this.Progress += new ProgressEventDelegate(this.CurrentExporter_Progress);
            this.ProgressMessage += new MessageEventDelegate(CurrentExporter_ProgressMessage);
            this.Aborted += new ErrorEventDelegate(this.CurrentExporter_Aborted);
            this.Finished += new BasicDelegate(this.CurrentExporter_Finished);
        }

        private void UnregisterExporterEvents()
        {
            try
            {
                this.Progress -= this.CurrentExporter_Progress;
                this.Aborted -= this.CurrentExporter_Aborted;
                this.Finished -= this.CurrentExporter_Finished;
                this.ProgressMessage -= CurrentExporter_ProgressMessage;
            }
            catch (Exception ex)
            {
                logger.Error($"UnregisterExporterEvents raise an error{ex.Message}");
            }
        }

        /// <summary>
        /// Stop export
        /// </summary>
        public void Stop()
        {
            lock (this.m_locker)
            {
                if (this.m_currentExporter != null)
                {
                    this.UnregisterExporterEvents();
                    this.Stop();
                    this.FreeCurrentExporter();
                }
            }
        }

        private void FreeCurrentExporter()
        {
            try
            {
                lock (this.m_locker)
                {
                    if (this.m_currentExporter != null)
                    {
                        this.UnregisterExporterEvents();

                        this.m_currentExporter.Dispose();
                        this.m_currentExporter = null;
                    }
                }
            }
            catch
            { }
        }

        /// <summary>
        /// Gets name of the file to export
        /// </summary>
        /// <param name="result">TestResults</param>
        /// <returns>Name of the file</returns>
        //protected virtual string GetFileName()
        //{
        //    return System.IO.Path.Combine(this.m_fileName, String.Format("{0} {1}, {2:dd-MMM-yyyy} {3:00}h{4:00}m{5:00}s.xlsx",
        //        "Application-Log", string.Empty, result.StartTime, result.StartTime.Hour, result.StartTime.Minute, result.StartTime.Second));
        //}

        private void OnProgress(double progress)
        {
            if (this.Progress != null)
                this.Progress(this, progress);
        }

        private void OnAbort(object error)
        {
            if (this.Aborted != null)
                this.Aborted(error);
        }

        private void OnFinish()
        {
            if (this.Finished != null)
                this.Finished();
        }

        private void CurrentExporter_Progress(object sender, double progress)
        {
            if (this.m_logList.Count > 0)
                progress = (this.m_currentInnerResultIdx + progress) / this.m_logList.Count;
            else
                progress = 1;

            this.OnProgress(progress);
        }

        private void CurrentExporter_ProgressMessage(string message)
        {
            if (this.ProgressMessage != null)
                this.ProgressMessage(message);
        }

        private void CurrentExporter_Aborted(object error)
        {
            this.FreeCurrentExporter();
            this.OnAbort(error);
        }

        private void CurrentExporter_Finished()
        {
            try
            {
                //if (!this.ProcessNextInnerResult())
                //    this.OnFinish();
            }
            catch (Exception e)
            {
                this.FreeCurrentExporter();
                this.OnAbort(e);
            }
        }

        private void OnExportAborted(Exception e)
        {
            if (this.Aborted != null)
                this.Aborted(e);
        }

        private void OnFinished()
        {
            if (this.Finished != null)
                this.Finished();
        }

        private void ReportProgress(string message)
        {
            if (this.ProgressMessage != null)
                this.ProgressMessage(message);
        }

        #region IDisposable Members

        /// <summary>
        /// Clean up any resources being used
        /// </summary>
        public void Dispose()
        {
            lock (this.m_locker)
            {
                if (this.m_currentExporter != null)
                {
                    if (_ref == 1)
                    {
                        if (_builder != null)
                            _builder = null;
                        Interlocked.Decrement(ref _ref);
                    }
                    this.Dispose();
                }
            }
        }

        #endregion
    }
}
