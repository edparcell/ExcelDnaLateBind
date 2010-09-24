using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Interior : LateBoundObject
    {
        public Interior(object Interior) : base(Interior) { }

        //Color 
        public int Color
        {
            get { return Convert.ToInt32((double)GetProperty("Color")); }
            set { SetProperty("Color", value); }
        }
        public string ColorRRGGBB { get { return ColorUtils.ColorToColorRRGGBB(Color); } }

        //ColorIndex 
        public int ColorIndex
        {
            get { return (int)GetProperty("ColorIndex"); }
            set { SetProperty("ColorIndex", value); }
        }
    }
}
