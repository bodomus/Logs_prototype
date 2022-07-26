using System;
using System.Collections.Generic;
using System.Text;

namespace PrototypeLogs.Export
{
    public class LogItem
    {
        /// <summary>
        /// Name operation
        /// </summary>
        public string Action { get; set; }
        /// <summary>
        /// value of the Action ie action-press button value - left
        /// </summary>
        public string Value { get; set; }
        /// <summary>
        /// description
        /// </summary>
        public string Description { get; set; }
       
    }
}
