using ExcelLibrary.Infrastructure;

namespace ExcelLibrary.BinaryFileFormat
{
    public partial class FONT : Record
    {
        public byte UnderlineStyle
        {
            get { return Underline; }
            set
            {
                OptionFlags = OptionFlags.SetBit(value != 0, 2);
                Underline = value;
            }
        }

        public bool Bold
        {
            get { return OptionFlags.GetBit(0); }
            set
            {
                Weight = (ushort) (value ? 700 : 400);
                OptionFlags = OptionFlags.SetBit(value, 0);
            }
        }

        public bool Italic
        {
            get { return OptionFlags.GetBit(1); }
            set { OptionFlags = OptionFlags.SetBit(value, 1); }
        }

        public bool Strikethrough
        {
            get { return OptionFlags.GetBit(3); }
            set { OptionFlags = OptionFlags.SetBit(value, 3); }
        }

        public bool Outlined
        {
            get { return OptionFlags.GetBit(4); }
            set { OptionFlags = OptionFlags.SetBit(value, 4); }
        }

        public bool Shadowed
        {
            get { return OptionFlags.GetBit(5); }
            set { OptionFlags = OptionFlags.SetBit(value, 5); }
        }

        public bool Condensed
        {
            get { return OptionFlags.GetBit(6); }
            set { OptionFlags = OptionFlags.SetBit(value, 6); }
        }

        public bool Extended
        {
            get { return OptionFlags.GetBit(7); }
            set { OptionFlags = OptionFlags.SetBit(value, 7); }
        }
    }
}