using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Automation.Peers;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ColorChat.WPF.EventLogger;
using NLog;

namespace Pathway.WPF.ImportExport
{
    public class EventLoggerClass : IEventLoggerClass
    {
        private const int MAX_LOG_BOUND = 10;

        private static Logger eventLog = LogManager.GetLogger("eventLog");
        private int _countProperties;
        private int? _timestamp;
        private Point? _lastMouseClickPoint;
        private FrameworkElement CurrentFocus;

        /// <summary>
        /// key is number of order
        /// </summary>
        private static Dictionary<int, LogItem> _properties = new Dictionary<int, LogItem>();

        public EventLoggerClass()
        {
            _countProperties = 0;
            EventManager.RegisterClassHandler(
                typeof(Control),
                UIElement.MouseDownEvent,
                new MouseButtonEventHandler(MouseDown),
                true
            );

            EventManager.RegisterClassHandler(
                typeof(Control),
                UIElement.KeyDownEvent,
                new KeyEventHandler(KeyDown),
                true
            );

            EventManager.RegisterClassHandler(
                typeof(UIElement),
                UIElement.PreviewTextInputEvent,
                new TextCompositionEventHandler(TextInput),
                true
            );
        }

        private void ClearProperties()
        {
            _properties?.Clear();
            _countProperties = 0;
        }

        private void AddLogItem(LogItem item, bool disableAutoReset)
        {
            eventLog.Info($"{item.Action} {item.Value} {item.Description}");
        }

        /// <summary>
        /// Action
        /// </summary>
        /// <param name="source"></param>
        private LogItem CollectCommonProperties(TypeAction @type, FrameworkElement source, bool disableAutoReset)
        {
            LogItem li = new LogItem { Action = $"A: {@type.ToString()}", Value = $"V: {source.DependencyObjectType.Name}", Description = "D:" };
            return li;
        }

        void MouseDown(object sender, MouseButtonEventArgs e)
        {
            FrameworkElement source = sender as FrameworkElement;
            if (source == null)
                return;

            Point pt = e.GetPosition((UIElement)e.OriginalSource);
            if (!_lastMouseClickPoint.HasValue || !(pt.X == _lastMouseClickPoint.Value.X && pt.Y == _lastMouseClickPoint.Value.Y))
                _lastMouseClickPoint = new Point { X = pt.X, Y = pt.Y };
            else
            {
                if (pt.X == _lastMouseClickPoint.Value.X && pt.Y == _lastMouseClickPoint.Value.Y)
                    return;
            }
            // Perform the hit test against a given portion of the visual object tree.
            HitTestResult result = VisualTreeHelper.HitTest(source, pt);
            var li = CollectCommonProperties(TypeAction.MousePress, e.OriginalSource as FrameworkElement, true);

            LogMouse(li, e, pt);
        }

        void LogMouse(LogItem li, MouseButtonEventArgs e, Point pt)
        {
            string result = "";
            if (e.LeftButton == MouseButtonState.Pressed)
                result += "Left button pressed ";
            if (e.RightButton == MouseButtonState.Pressed)
                result += " Right button pressed ";
            if (e.MiddleButton == MouseButtonState.Pressed)
                result += " Middle button pressed ";

            if (e.ClickCount == 2)
            {
                li.Description = $"D: Double click " + result;
            }
            else
            {
                li.Description = $"D: " + result;
            }

            li.Action += $"X:{(int)pt.X} Y:{(int)pt.Y}";
            AddLogItem(li, false);
        }

        void KeyDown(object sender, KeyEventArgs e)
        {
            FrameworkElement source = sender as FrameworkElement;
            if (source == null)
                return;
            if (!source.Focusable || source.GetType().Name == "MainWindow" || source.GetType().Name == "ScrollViewer")
                return;
            LogKeyboard(e, source, isUp: false);
        }

        void LogKeyboard(KeyEventArgs e, FrameworkElement source, bool isUp)
        {
            var action = $"A:KP: {source.GetType().Name}";
            var kv = $"V: {e.Key.ToString()}";
            var Description = $"D: Modifier: {Keyboard.Modifiers} SystemKey: {e.SystemKey.ToString()}";
            var item = new LogItem
            {
                Action = action,
                Value = kv,
                Description = (Keyboard.Modifiers != ModifierKeys.None || e.SystemKey != Key.None) ? Description : "D:"
            };
            AddLogItem(item, false);
        }

        bool CheckPasswordElement(UIElement targetElement)
        {
            if (targetElement != null)
            {
                AutomationPeer automationPeer = UIElementAutomationPeer.CreatePeerForElement(targetElement);
                return (automationPeer != null) ? automationPeer.IsPassword() : false;
            }
            return false;
        }

        //Key GetKeyValue(Key key)
        //{
        //    return key switch
        //    {
        //        Key.Tab or Key.Left or Key.Right or Key.Up or Key.Down or Key.PageUp or
        //        Key.PageDown or Key.LeftCtrl or Key.RightCtrl or Key.LeftShift or Key.RightShift or
        //        Key.Enter or Key.Home or Key.End => key,
        //        _ => key,
        //    };
        //}

        public void Print()
        {
            foreach (var item in _properties)
            {
                var v = item.Value;
                eventLog.Info($"{v.Action} {v.Value} {v.Description}");
            }
        }

        void TextInput(object sender, TextCompositionEventArgs e)
        {
            FrameworkElement source = sender as FrameworkElement;
            if (source == null)
                return;
            LogTextInput(e);
        }

        void LogTextInput(TextCompositionEventArgs e)
        {
            var element = e.OriginalSource as UIElement;
            var el = element.GetType().Name;
            if (!_timestamp.HasValue || _timestamp.Value != e.Timestamp)
            {
                _timestamp = e.Timestamp;
                AddLogItem(new LogItem { Action = $"A:TI {el} ", Value = $"V: {e.Text}", Description = $"D:" }, false);
            }
        }

    }
}
