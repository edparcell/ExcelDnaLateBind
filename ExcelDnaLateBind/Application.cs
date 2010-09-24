using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using System.Text.RegularExpressions;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Application : LateBoundObject
    {
        public Application(object Application) : base(Application) { }

        public static Application Default { get; private set;  }
        static Application()
        {
            Default = new Application(ExcelDnaUtil.Application);
        }

        public Workbook ActiveWorkbook
        {
            get
            {
                return new Workbook(GetProperty("ActiveWorkbook"));
            }
        }

        public Workbooks Workbooks
        {
            get
            {
                return new Workbooks(GetProperty("Workbooks"));
            }
        }

        public Worksheet ActiveWorksheet {
            get
            {
                return new Worksheet(GetProperty("ActiveSheet"));
            }
        }

        public Range Range(ExcelReference reference)
        {
            string internalSheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            Match match = Regex.Match(internalSheetName, @"\[(.*)\](.*)");
            string workbookName = match.Groups[1].Value;
            string sheetName = match.Groups[2].Value;
            Range TopLeft = Workbooks[workbookName].Sheets[sheetName].Cells(reference.RowFirst + 1, reference.ColumnFirst + 1);
            Range BottowRight = Workbooks[workbookName].Sheets[sheetName].Cells(reference.RowLast + 1, reference.ColumnLast + 1);
            Range rng = Range(TopLeft, BottowRight);
            return rng;
        }

        public Range Range(Range rng1, Range rng2)
        {
            return new Range(GetProperty("Range", new object[] { rng1.Obj, rng2.Obj }));
        }
    }
}
