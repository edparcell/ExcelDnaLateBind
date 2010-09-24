using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Range : LateBoundObject
    {
        public Range(object Range) : base(Range) { }

        public object Value
        {
            get
            {
                return GetProperty("Value");
            }
            set
            {
                SetProperty("Value", value);
            }
        }
    }
}
