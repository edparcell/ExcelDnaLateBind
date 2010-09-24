using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CobaltQuantware.ExcelDnaLateBind
{
    public class Font : LateBoundObject
    {
        public Font(object Font) : base(Font) { }

        //Application
        //Background 
        //Bold 
        public bool Bold
        {
            get { return (bool)GetProperty("Bold"); }
            set { SetProperty("Bold", value); }
        }

        //Color 
        public int Color { 
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

        //FontStyle 
        //Italic 
        public bool Italic
        {
            get { return (bool)GetProperty("Italic"); }
            set { SetProperty("Italic", value); }
        }
        
        //Name
        public string Name {
            get { return (string)GetProperty("Name"); }
            set { SetProperty("Name", value); }
        }

        //Parent 
        
        //Size 
        public int Size
        {
            get { return (int)GetProperty("Size"); }
            set { SetProperty("Size", value); }
        }

        //Strikethrough 
        //Subscript 
        //Superscript 
        //ThemeColor 
        //ThemeFont 
        //TintAndShade 
        //Underline 

    }
}