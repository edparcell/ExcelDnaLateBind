using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Workbook : LateBoundObject
    {
        public Workbook(object Workbook) : base(Workbook) { }

        public Worksheets Sheets { get { return new Worksheets(GetProperty("Sheets")); } }
    }
}
