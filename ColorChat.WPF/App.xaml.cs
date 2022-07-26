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
        private static Logger logger1 = LogManager.GetLogger("eventFile");
        private EventLoggerClass EL = new EventLoggerClass();
        protected override void OnExit(ExitEventArgs e)
        {
            EL.Print();
            base.OnExit(e);
        }
        protected override void OnStartup(StartupEventArgs e)
        {
            logger1.Info("Note \n A: \tAction\n V: \tValue\n D: \tDesciption\n KP: \tKeyPress\n TI: \tText input\n MP: \tMouse press\n ");
           
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
