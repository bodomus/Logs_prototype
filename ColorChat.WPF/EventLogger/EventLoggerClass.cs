using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Automation.Peers;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

using NLog;

namespace ColorChat.WPF.EventLogger
{
    public class EventLoggerClass : IEventLoggerClass
    {
        private const int MAX_LOG_BOUND = 10;

        private static Logger eventFile = LogManager.GetLogger("eventFile");
        private int _countProperties;
        private int? _timestamp;
        private Point? _lastMouseClickPoint;
        private FrameworkElement CurrentFocus;

        /// <summary>
        /// key is number of order
        /// </summary>
        private Dictionary<int, LogItem> _properties = new Dictionary<int, LogItem>();

        public EventLoggerClass()
        {
            _countProperties = 0;
            eventFile.Info("Start collect events...");
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
            _properties.Add(++_countProperties, item);
            if (_countProperties > MAX_LOG_BOUND && !disableAutoReset)
            {
                Print();
                ClearProperties();
            }
        }

        /// <summary>
        /// Action
        /// </summary>
        /// <param name="source"></param>
        private int CollectCommonProperties(TypeAction @type, FrameworkElement source, bool disableAutoReset)
        {
            AddLogItem(new LogItem { Action = $"A:{@type.ToString()}", Value = $"V: {source.DependencyObjectType.Name}" }, disableAutoReset);
            return _countProperties;
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
            int key = CollectCommonProperties(TypeAction.MousePress, e.OriginalSource as FrameworkElement, true);
            LogMouse(key, e, isUp: false);
        }

        void LogMouse(int key, MouseButtonEventArgs e, bool isUp)
        {
            if (!_properties.TryGetValue(key, out var logItem))
            {
                throw new ArgumentException($"Log mouse: invalid key: {key} in log dictionary.");
            }

            if (e.ClickCount == 2)
            {
                logItem.Description = "doubleClick";
            }
            else if (isUp)
            {
                logItem.Description = "up";
            }
            else
            {
                logItem.Description = "Down";
            }
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
            var kv = $"V: {GetKeyValue(e.Key).ToString()}";
            var Description = $"D: Modifier: {Keyboard.Modifiers} SystemKey: {e.SystemKey.ToString()}";
            var item = new LogItem
            {
                Action = action,
                Value = kv,
                Description = (Keyboard.Modifiers != ModifierKeys.None || e.SystemKey != Key.None) ? Description : string.Empty
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

        Key GetKeyValue(Key key)
        {
            return key switch
            {
                Key.Tab or Key.Left or Key.Right or Key.Up or Key.Down or Key.PageUp or
                Key.PageDown or Key.LeftCtrl or Key.RightCtrl or Key.LeftShift or Key.RightShift or
                Key.Enter or Key.Home or Key.End => key,
                _ => key,
            };
        }

        public void Print()
        {
            foreach (var item in _properties)
            {
                var v = item.Value;
                eventFile.Info($"{v.Action} {v.Value} {v.Description}");
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
                AddLogItem(new LogItem { Action = $"A:TI {el} TimeStamp: {e.Timestamp}", Value = $"V: {e.Text}" }, false);
            }
        }

        private string GetDescription(MouseButtonEventArgs e)
        {
            string objName = string.Empty;
            if (e.OriginalSource is System.Windows.Controls.TextBlock && e.Source is FrameworkElement)
                objName = (e.OriginalSource as System.Windows.Controls.TextBlock).Text;
            //if (e.OriginalSource is System.Windows.Controls.Image && e.Source is FrameworkElement)
            //    objName = this.GetObjectNameByImageName((e.OriginalSource as System.Windows.Controls.Image).Source.ToString());
            if (FindUpVisualTree<System.Windows.Controls.TextBox>(e.OriginalSource as DependencyObject) != null && e.Source is FrameworkElement)
                objName = (FindUpVisualTree<System.Windows.Controls.TextBox>(e.OriginalSource as DependencyObject)).Text;

            if (!string.IsNullOrWhiteSpace(objName))
                return string.Format("Click on '{0}', '{1}'", objName, (e.Source as FrameworkElement).Name);
            else
                return string.Empty;
        }

        public static T FindUpVisualTree<T>(DependencyObject initial) where T : DependencyObject
        {
            DependencyObject current = initial;

            while (current != null && current.GetType() != typeof(T))
            {
                current = VisualTreeHelper.GetParent(current);
            }
            return current as T;
        }
    }
}
