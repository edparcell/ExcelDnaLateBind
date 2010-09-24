using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Workbooks : LateBoundCollection<Workbook>
    {
        public Workbooks(object Workbooks) : base(Workbooks) { }
    }
}