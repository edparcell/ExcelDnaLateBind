using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Worksheets : LateBoundCollection<Worksheet>
    {
        public Worksheets(object Worksheets) : base(Worksheets) { }
    }
}
