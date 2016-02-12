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
            get { return (byte)(LineStyle & 0x0F); }
            set
            {
                if (value > 0x0D)
                {
                    throw new ArgumentOutOfRangeException("LeftLineStyle");
                }
                LineStyle = LineStyle & 0xFFFFFFF0 | value;
            }
        }

        public int LeftLineColorIndex
        {
            get { return (int) ((LineStyle & 0x7F0000) >> 16); }
            set
            {
                LineStyle = LineStyle & 0xFFFFFF0F | (uint)value << 4;
            }
        }

        public byte RightLineStyle
        {
            get { return (byte)((LineStyle & 0xF0) >> 4); }
            set
            {
                if (value > 0x0D)
                {
                    throw new ArgumentOutOfRangeException("RightLineStyle");
                }
                LineStyle = LineStyle & 0xFFFFFF0F | (uint)value << 4;
            }
        }

        public byte TopLineStyle
        {
            get { return (byte)((LineStyle & 0xF00) >> 8); }
            set
            {
                if (value > 0x0D)
                {
                    throw new ArgumentOutOfRangeException("TopLineStyle");
                }
                LineStyle = LineStyle & 0xFFFFF0FF | (uint)value << 8;
            }
        }

        public byte BottomLineStyle
        {
            get { return (byte)((LineStyle & 0xF000) >> 12); }
            set
            {
                if (value > 0x0D)
                {
                    throw new ArgumentOutOfRangeException("BottomLineStyle");
                }
                LineStyle = LineStyle & 0xFFFF0FFF | (uint)value << 12;
            }
        }
    }
}
