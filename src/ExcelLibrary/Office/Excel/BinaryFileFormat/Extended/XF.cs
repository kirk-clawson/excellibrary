using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ExcelLibrary.Infrastructure;

namespace ExcelLibrary.BinaryFileFormat
{
    public partial class XF : Record
    {
        public int PatternColorIndex
        {
            get { return Background.GetBits(0, 7); }
            set
            {
                if (value > 0x7F)
                {
                    throw new ArgumentOutOfRangeException("PatternColorIndex");
                }
                Background = Background.SetBits((ushort)value, 0, 7);
            }
        }

        public int PatternBackgroundColorIndex
        {
            get { return Background.GetBits(7, 7); }
            set
            {
                Background = Background.SetBits((ushort)value, 7, 7);
            }
        }

        public int FillPattern
        {
            get { return (int)LineColor.GetBits(26, 6); }
            set
            {
                LineColor = LineColor.SetBits((uint)value, 26, 6);
            }
        }

        public byte LeftLineStyle
        {
            get { return (byte)LineStyle.GetBits(0, 4); }
            set
            {
                if (value > 0x0D)
                {
                    throw new ArgumentOutOfRangeException("LeftLineStyle");
                }
                LineStyle = LineStyle.SetBits(value, 0, 4);
            }
        }

        public int LeftLineColorIndex
        {
            get { return (int) LineStyle.GetBits(16, 7); }
            set
            {
                LineStyle = LineStyle.SetBits((uint)value, 16, 7);
            }
        }

        public byte RightLineStyle
        {
            get { return (byte)LineStyle.GetBits(4, 4); }
            set
            {
                if (value > 0x0D)
                {
                    throw new ArgumentOutOfRangeException("RightLineStyle");
                }
                LineStyle = LineStyle.SetBits(value, 4, 4);
            }
        }

        public int RightLineColorIndex
        {
            get { return (int)LineStyle.GetBits(23, 7); }
            set
            {
                LineStyle = LineStyle.SetBits((uint)value, 23, 7);
            }
        }

        public byte TopLineStyle
        {
            get { return (byte)LineStyle.GetBits(8, 4); }
            set
            {
                if (value > 0x0D)
                {
                    throw new ArgumentOutOfRangeException("TopLineStyle");
                }
                LineStyle = LineStyle.SetBits(value, 8, 4);
            }
        }

        public int TopLineColorIndex
        {
            get { return (int)LineColor.GetBits(0, 7); }
            set
            {
                LineColor = LineColor.SetBits((uint)value, 0, 7);
            }
        }

        public byte BottomLineStyle
        {
            get { return (byte)LineStyle.GetBits(12, 4); }
            set
            {
                if (value > 0x0D)
                {
                    throw new ArgumentOutOfRangeException("BottomLineStyle");
                }
                LineStyle = LineStyle.SetBits(value, 12, 4);
            }
        }

        public int BottomLineColorIndex
        {
            get { return (int)LineColor.GetBits(7, 7); }
            set
            {
                LineColor = LineColor.SetBits((uint)value, 7, 7);
            }
        }

        public byte DiagonalLineStyle
        {
            get { return (byte)LineColor.GetBits(21, 4); }
            set
            {
                if (value > 0x0D)
                {
                    throw new ArgumentOutOfRangeException("DiagonalLineStyle");
                }
                LineColor = LineColor.SetBits(value, 21, 4);
            }
        }

        public int DiagonalLineColorIndex
        {
            get { return (int)LineColor.GetBits(14, 7); }
            set
            {
                LineColor = LineColor.SetBits((uint)value, 14, 7);
            }
        }

        public bool DiagonalDown
        {
            get { return LineStyle.GetBit(30); }
            set { LineStyle = LineStyle.SetBit(value, 30); }
        }

        public bool DiagonalUp
        {
            get { return LineStyle.GetBit(31); }
            set { LineStyle = LineStyle.SetBit(value, 31); }
        }

        public bool CellLocked
        {
            get { return CellProtection.GetBit(0); }
            set { CellProtection = CellProtection.SetBit(value, 0); }
        }

        public bool FormulaHidden
        {
            get { return CellProtection.GetBit(1); }
            set { CellProtection = CellProtection.SetBit(value, 1); }
        }

        public byte HorizontalAlign
        {
            get { return Alignment.GetBits(0, 3); }
            set { Alignment = Alignment.SetBits(value, 0, 3); }
        }

        public byte VerticalAlign
        {
            get { return Alignment.GetBits(4, 3); }
            set { Alignment = Alignment.SetBits(value, 4, 3); }
        }

        public bool TextWrap
        {
            get { return Alignment.GetBit(3); }
            set { Alignment = Alignment.SetBit(value, 3); }
        }

        public bool JustifyDistributed
        {
            get { return Alignment.GetBit(7); }
            set { Alignment = Alignment.SetBit(value, 7); }
        }

        public byte IndentLevel
        {
            get { return Indent.GetBits(0, 4); }
            set { Indent = Indent.SetBits(value, 0, 4); }
        }

        public bool ShrinkContents
        {
            get { return Indent.GetBit(4); }
            set { Indent = Indent.SetBit(value, 4); }
        }

        public byte TextDirection
        {
            get { return Indent.GetBits(6, 2); }
            set { Indent = Indent.SetBits(value, 6, 2); }
        }
    }
}
