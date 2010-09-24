using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    internal class ColorUtils
    {
        internal static string ColorToColorRRGGBB(int color)
        {
            int r = color % 256;
            int g = (color / 256) % 256;
            int b = (color / 256 / 256) % 256;
            color = 256 * 256 * r + 256 * g + b;
            return color.ToString("X6");
        }
    }
}
