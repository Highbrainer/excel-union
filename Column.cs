using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUnion
{
    class Column
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public int index { get; set; }

        public Column(int index, string title)
        {
            this.index = index;
            this.Title = title;
            this.Name = ""+(char)(((int)'A') + index);
        }
        public override string ToString()
        {
            return Title;
        }
    }
}
