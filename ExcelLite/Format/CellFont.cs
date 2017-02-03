using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLite.Format
{
    class CellFont
    {
        private string _fontName;
        public string FontName
        {
            get { return _fontName; }
            set { _fontName = value; }
        }

        private double _fontSize;
        public double FontSize
        {
            get { return _fontSize; }
            set { _fontSize = value; }
        }

        private string _fontFamily;
        public string FontFamily
        {
            get { return _fontFamily; }
            set { _fontFamily = value; }
        }

        private string _fontColor;
        public string FontColor
        {
            get { return _fontColor; }
            set { _fontColor = value; }
        }

        private bool _bold;
        public bool Bold
        {
            get { return _bold; }
            set { _bold = value; }
        }

        public bool _italic;
        public bool Italic
        {
            get { return _italic; }
            set { _italic = value; }
        }

        private bool _underLine;
        public bool UnderLine
        {
            get { return _underLine; }
            set { _underLine = value; }
        }

        public CellFont()
        {
            
        }
    }
}
