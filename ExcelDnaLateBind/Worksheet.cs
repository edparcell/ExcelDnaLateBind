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
    }
}
