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
    public class EventLoggerClass: IEventLoggerClass
    {
        private static Logger logger = LogManager.GetLogger("file");
        private static Logger logger1 = LogManager.GetLogger("file1");
        private int _countProperties;
        private Point? _lastMouseClickPoint;
        private int? _timestamp;
        private FrameworkElement CurrentFocus;
        /// <summary>
        /// key is number of order
        /// </summary>
        private Dictionary<int, LogItem> _properties = new Dictionary<int, LogItem>();

        public EventLoggerClass()
        {
            _countProperties = 0;
            logger.Info("Start EventLoggerClass");
            logger1.Info("Start EventLoggerClass");
            //EventManager.RegisterClassHandler(
            //    typeof(Control),
            //    UIElement.MouseDownEvent,
            //    new MouseButtonEventHandler(MouseDown),
            //    true
            //);

            EventManager.RegisterClassHandler(
                typeof(Control),
                UIElement.KeyDownEvent,
                new KeyEventHandler(KeyDown),
                true
            );

            EventManager.RegisterClassHandler(
                typeof(Control),
                Keyboard.GotKeyboardFocusEvent,
                new KeyboardFocusChangedEventHandler(OnKeyboardFocusChanged),
                true
            );

            EventManager.RegisterClassHandler(
                typeof(UIElement),
                UIElement.PreviewTextInputEvent,
                new TextCompositionEventHandler(TextInput),
                true
            );
        }

        /// <summary>
        /// Action
        /// </summary>
        /// <param name="source"></param>
        private int CollectCommonProperties(TypeAction @type, FrameworkElement source)
        {
            _properties.Add(++_countProperties, new LogItem { Action = $"A:{@type.ToString()}", Value = $"V: {source.DependencyObjectType.Name}"});
            return _countProperties;
        }

        void MouseDown(object sender, MouseButtonEventArgs e)
        {
            FrameworkElement source = sender as FrameworkElement;
            Control source1 = sender as Control;
            if (source == null)
                return;

            Point pt = e.GetPosition((UIElement)e.OriginalSource);
            if (!_lastMouseClickPoint.HasValue || !(pt.X == _lastMouseClickPoint.Value.X && pt.Y == _lastMouseClickPoint.Value.Y))
                _lastMouseClickPoint = new Point {X=pt.X, Y= pt.Y };
            else {
                if (pt.X == _lastMouseClickPoint.Value.X && pt.Y == _lastMouseClickPoint.Value.Y)
                {
                    return;
                }
            }
            // Perform the hit test against a given portion of the visual object tree.
            HitTestResult result = VisualTreeHelper.HitTest(source, pt);
            //var b = result.VisualHit.GetValue();
            if (result != null)
            {
                // Perform action on hit visual object.
            }
            int key = CollectCommonProperties(TypeAction.MousePress, e.OriginalSource as FrameworkElement);
            LogMouse(key, _properties, e, isUp: false);
        }

        void LogMouse(int key, IDictionary<int, LogItem> properties, MouseButtonEventArgs e, bool isUp)
        {
            if (!properties.TryGetValue(key, out var logItem)){
                throw new Exception("Invalid key");
            }

            //logItem.Value = e.ChangedButton.ToString();
            //logItem.Description = e.ClickCount.ToString();
            //properties["mouseButton"] = e.ChangedButton.ToString();
            //properties["ClickCount"] = e.ClickCount.ToString();
            //Breadcrumb item = new Breadcrumb();
            if (e.ClickCount == 2)
            {
                //properties["action"] = "doubleClick";
                logItem.Description = "doubleClick";
                //item.Event = BreadcrumbEvent.MouseDoubleClick;
            }
            else if (isUp)
            {
                //properties["action"] = "up";
                logItem.Description = "up";
                //item.Event = BreadcrumbEvent.MouseUp;
            }
            else
            {
                //properties["action"] = "down";
                logItem.Description = "Down";
                //item.Event = BreadcrumbEvent.MouseDown;
            }

            //_properties.Add(++_countProperties, logItem);
            //item.CustomData = properties;

            //AddBreadcrumb(item);
        }

        void KeyDown(object sender, KeyEventArgs e)
        {
            FrameworkElement source = sender as FrameworkElement;
            if (source == null)
                return;
            if (!source.Focusable || source.GetType().Name == "MainWindow" || source.GetType().Name == "ScrollViewer")
                return;
            LogKeyboard(_properties, e, source, isUp: false);
        }

        void LogKeyboard(IDictionary<int, LogItem> properties, KeyEventArgs e, FrameworkElement source, bool isUp)
        {
            var action = $"A:KP: {source.GetType().Name}";
            var value = $"V: {e.Key}";
            var Description = $"D: Modifier: {Keyboard.Modifiers} SystemKey: {e.SystemKey.ToString()}";
            var kv = GetKeyValue(e.Key).ToString();
            var item = new LogItem { Action = action, Value = kv, 
                Description = (Keyboard.Modifiers != ModifierKeys.None || e.SystemKey != Key.None)? Description: string.Empty };
            _properties.Add(++_countProperties, item);
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
            switch (key)
            {
                case Key.Tab:
                case Key.Left:
                case Key.Right:
                case Key.Up:
                case Key.Down:
                case Key.PageUp:
                case Key.PageDown:
                case Key.LeftCtrl:
                case Key.RightCtrl:
                case Key.LeftShift:
                case Key.RightShift:
                case Key.Enter:
                case Key.Home:
                case Key.End:
                    return key;

                default:
                    return key;
            }
        }

        public void Print()
        {
            foreach (var item in _properties)
            {
                var i = item.Key;
                var v = item.Value;
                logger.Info($"logger: #{i} {v.Action} {v.Value} {v.Description}");
                logger1.Info($"logger1: #{i} {v.Action} {v.Value} {v.Description}");
            }
        }

        void TextInput(object sender, TextCompositionEventArgs e)
        {
            FrameworkElement source = sender as FrameworkElement;
            if (source == null)
                return;
            CheckPasswordElement(source);
            LogTextInput(_properties, e);
        }

        void LogTextInput(IDictionary<int, LogItem> properties, TextCompositionEventArgs e)
        {
            var element = e.OriginalSource as UIElement;
            var el = element.GetType().Name;
            if (!_timestamp.HasValue || _timestamp.Value != e.Timestamp)
            {
                _timestamp = e.Timestamp;
                _properties.Add(++_countProperties, new LogItem { Action = $"A:TI {el} TimeStamp {e.Timestamp}", Value = $"V: {e.Text}" });
            }
        }

        void OnKeyboardFocusChanged(object sender, KeyboardFocusChangedEventArgs e)
        {
            FrameworkElement oldFocus = e.OldFocus as FrameworkElement;
            if (oldFocus != null)
            {
                //LogFocus(isGotFocus: false, e);
            }

            FrameworkElement newFocus = e.NewFocus as FrameworkElement;
            if (newFocus != null)
            {
                CurrentFocus = sender as FrameworkElement;
                //LogFocus(isGotFocus: true, e);
            }
        }

        void LogFocus(bool isGotFocus, KeyboardFocusChangedEventArgs e)
        {
            _properties.Add(++_countProperties, new LogItem { Action = $"A:{(isGotFocus == false ? "LostFocus" : "GotFocus")}", Value = $"V: {e.OriginalSource.ToString()}" });
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
