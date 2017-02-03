using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLite.Format
{
    public class CellAlignment
    {
        private HorizonalAlignment _horizonal;
        public HorizonalAlignment Horizonal
        {
            get { return _horizonal; }
            set { _horizonal = value; }
        }

        public VerticalAlignment _vertical;
        public VerticalAlignment Vertical
        {
            get { return _vertical; }
            set { _vertical = value; }
        }

        private bool _wrapText;
        public bool WrapText
        {
            get { return _wrapText; }
            set { _wrapText = value; }
        }

        public CellAlignment()
        {
            _horizonal = HorizonalAlignment.Left;
            _vertical = VerticalAlignment.Top;
            _wrapText = true;
        }
    }

    public enum HorizonalAlignment
    {
        Left,
        Center,
        Right
    }

    public enum VerticalAlignment
    {
        Top,
        Center,
        Bottom
    }
}
