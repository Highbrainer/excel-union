using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUnion
{
    class Content
    {
        public string Sheet { get; set; }
        public Dictionary<string, List<object>> Lines { get; set; } = new Dictionary<string, List<object>>();
        public List<object> Titles { get; set; } = new List<object>();
    }
}
