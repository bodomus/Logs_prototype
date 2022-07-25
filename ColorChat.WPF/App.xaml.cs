using ColorChat.WPF.Services;
using ColorChat.WPF.ViewModels;
using Microsoft.AspNetCore.SignalR.Client;
using NLog;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ColorChat.WPF.EventLogger;

namespace ColorChat.WPF
{
    public partial class App : Application
    {
        private static Logger logger = LogManager.GetLogger("file");
        private static Logger logger1 = LogManager.GetLogger("file1");
        private EventLoggerClass EL = new EventLoggerClass();
        protected override void OnExit(ExitEventArgs e)
        {
            EL.Print();
            base.OnExit(e);
        }
        protected override void OnStartup(StartupEventArgs e)
        {
            logger.Info("Info");
            logger.Info("Note \n A: - Action\n V: - Value\n D: Desciption\n KP: KeyPress\n TI: Text input\n ");
            //var logsFiles = LogsExporter.GetLogs();
            //var excelFile = LogsExporter.GetExcelFileName();
            //var logExport = new LogsExporter(logsFiles, excelFile);
            
            HubConnection connection = new HubConnectionBuilder()
                .WithUrl("http://localhost:5000/colorchat")
                .Build();

            ColorChatViewModel chatViewModel = ColorChatViewModel.CreatedConnectedViewModel(new SignalRChatService(connection));

            MainWindow window = new MainWindow
            {
                DataContext = new MainViewModel(chatViewModel)
            };

            window.Show();
        }
    }
}
