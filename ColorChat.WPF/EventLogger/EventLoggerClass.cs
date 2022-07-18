using NLog;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ColorChat.WPF.EventLogger
{
    public class EventLoggerClass: IEventLoggerClass
    {
        private static Logger logger = LogManager.GetLogger("file");
        private static Logger logger1 = LogManager.GetLogger("file1");
        private int _countProperties;
        private Point? _lastMouseClickPoint;
        /// <summary>
        /// key is number of order
        /// </summary>
        private Dictionary<int, LogItem> _properties = new Dictionary<int, LogItem>();

        public EventLoggerClass()
        {
            _countProperties = 0;
            
            EventManager.RegisterClassHandler(
                typeof(Control),
                UIElement.MouseDownEvent,
                new MouseButtonEventHandler(MouseDown),
                true
            );

            //EventManager.RegisterClassHandler(
            //    typeof(Control),
            //    UIElement.PreviewKeyDownEvent,
            //    new KeyEventHandler(KeyDown),
            //    true
            //);

            //EventManager.RegisterClassHandler(
            //    typeof(Control),
            //    Keyboard.GotKeyboardFocusEvent,
            //    new KeyboardFocusChangedEventHandler(OnKeyboardFocusChanged),
            //    true
            //);
        }

        /// <summary>
        /// Action
        /// </summary>
        /// <param name="source"></param>
        private int CollectCommonProperties(TypeAction @type, FrameworkElement source)
        {
            _properties.Add(++_countProperties, new LogItem { Action = @type.ToString(), Value = source.DependencyObjectType.Name});
            //_properties["Name"] = source.Name;
            //_properties["ClassName"] = source.GetType().ToString();
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

        void LogMouse(int key, IDictionary<int, LogItem> properties,
              MouseButtonEventArgs e,
              bool isUp)
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

            CollectCommonProperties(TypeAction.KeyboardPress, source);
            LogKeyboard(_properties, e.Key,
                        isUp: false);
        }

        void LogKeyboard(IDictionary<int, LogItem> properties,
                 Key key,
                 bool isUp)
        {
            _properties.Add(++_countProperties, new LogItem {Action = GetKeyValue(key).ToString() });
            //properties["key"] = GetKeyValue(key).ToString();
            //properties["action"] = isUp ? "up" : "down";

            //Breadcrumb item = new Breadcrumb();
            //item.Event = isUp ? BreadcrumbEvent.KeyUp : BreadcrumbEvent.KeyDown;
            //item.CustomData = properties;

            //AddBreadcrumb(item);
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
                    return Key.Multiply;
            }
        }

        public void Print()
        {
            foreach (var item in _properties)
            {
                var i = item.Key;
                var v = item.Value;
                logger.Info($"logger: #{i} : Action: {v.Action} Value: {v.Value} Desciption: {v.Description}");
                //logger1.Info($"logger1: #{i} : Action: {v.Action} Value: {v.Value} Desciption: {v.Description}");
            }
        }

        void TextInput(object sender, TextCompositionEventArgs e)
        {
            FrameworkElement source = sender as FrameworkElement;
            if (source == null)
                return;

            CollectCommonProperties(TypeAction.TextInput, source);
            LogTextInput(_properties, e);
        }

        void LogTextInput(IDictionary<int, LogItem> properties,
                          TextCompositionEventArgs e)
        {
            _properties.Add(++_countProperties, new LogItem { Action = "Press button", Value = e.Text });
            //properties["text"] = e.Text;
            //properties["action"] = "press";

            //Breadcrumb item = new Breadcrumb();
            //item.Event = BreadcrumbEvent.KeyPress;
            //item.CustomData = properties;

            //AddBreadcrumb(item);
        }

        void OnKeyboardFocusChanged(object sender, KeyboardFocusChangedEventArgs e)
        {
            FrameworkElement oldFocus = e.OldFocus as FrameworkElement;
            if (oldFocus != null)
            {
               
                LogFocus(isGotFocus: false, e);
            }

            FrameworkElement newFocus = e.NewFocus as FrameworkElement;
            if (newFocus != null)
            {
               
                LogFocus(isGotFocus: true, e);
            }
        }

        void LogFocus(bool isGotFocus, KeyboardFocusChangedEventArgs e)
        {
            _properties.Add(++_countProperties, new LogItem { Action = isGotFocus == false ? "LostFocus" : "GotFocus", Value = e.OriginalSource.ToString() });
            //Breadcrumb item = new Breadcrumb();
            //item.Event = isGotFocus ? BreadcrumbEvent.GotFocus :
            //                          BreadcrumbEvent.LostFocus;
            //item.CustomData = properties;

            //AddBreadcrumb(item);
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
