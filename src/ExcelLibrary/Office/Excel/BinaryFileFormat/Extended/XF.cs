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
    }
}
