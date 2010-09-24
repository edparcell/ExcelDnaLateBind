using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Application : LateBoundObject
    {
        public Application() : base(ExcelDnaUtil.Application) { }

        public Workbook ActiveWorkbook
        {
            get
            {
                return new Workbook(GetProperty("ActiveWorkbook"));
            }
        }

        public Worksheet ActiveWorksheet {
            get
            {
                return new Worksheet(GetProperty("ActiveSheet"));
            }
        }
    }
}
