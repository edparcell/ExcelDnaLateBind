using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Range : LateBoundObject
    {
        public Range(object Range) : base(Range) { }

        public object Value { get { return GetProperty("Value"); } set { SetProperty("Value", value); } }
        public object Text { get { return GetProperty("Text"); } }

        public Interior Interior { get { return new Interior(GetProperty("Interior")); } }

        public Font Font { get { return new Font(GetProperty("Font")); } }

        public Range Rows { get { return new Range(GetProperty("Rows")); } }
        public Range Columns { get { return new Range(GetProperty("Columns")); } }

        public int Count { get { return (int)GetProperty("Count"); } }

        public Range Cells(int Row, int Column) { return new Range(GetProperty("Cells", new object[] { Row, Column })); }
    }
}
