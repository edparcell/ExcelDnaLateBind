using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Worksheet : LateBoundObject
    {
        public Worksheet(object Worksheet) : base(Worksheet) { }

        public void Calculate()
        {
            InvokeMethod("Calculate");
        }

        public Range Range(string cell)
        {
            return new Range(GetProperty("Range", new object[] { cell }));
        }

        public Range Cells(int Row, int Column)
        {
            return new Range(GetProperty("Cells", new object[] { Row, Column }));
        }
    }
}
