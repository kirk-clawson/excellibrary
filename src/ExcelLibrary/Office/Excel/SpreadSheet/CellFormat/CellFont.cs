using ExcelLibrary.Infrastructure;

namespace ExcelLibrary.SpreadSheet
{
    public enum FontUnderlineStyle : byte
    {
        None = 0,
        Single = 1,
        Double = 2,
        SingleAccounting = 0x21,
        DoubleAccounting = 0x22
    }

    public enum FontFamilyType : byte
    {
        None = 0,
        Roman = 1,
        Swiss = 2,
        Modern = 3,
        Script = 4,
        Decorative = 5
    }

    public enum FontEscapementType : byte
    {
        None = 0,
        Superscript = 1,
        Subscript = 2
    }

    public class CellFont
    {
        public CellFont()
        {
            Height = 220; // 11pt
            Name = "Calibri";
            Color = ExcelColor.WindowTextForFonts;
            UnderlineStyle = FontUnderlineStyle.None;
            Family = FontFamilyType.Swiss;
            Escapement = FontEscapementType.None;
        }
        public short Height { get; set; }
        public string Name { get; set; }
        public ExcelColor Color { get; set; }
        public FontUnderlineStyle UnderlineStyle { get; set; }
        public FontFamilyType Family { get; set; }
        public FontEscapementType Escapement { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Strikethrough { get; set; }
        public bool Outlined { get; set; }
        public bool Shadowed { get; set; }
        public bool Condensed { get; set; }
        public bool Extended { get; set; }

        #region Overrides of Object

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns>
        /// A string that represents the current object.
        /// </returns>
        public override string ToString()
        {
            byte flags = 0;
            flags = flags.SetBit(Bold, 0);
            flags = flags.SetBit(Italic, 1);
            flags = flags.SetBit(Strikethrough, 3);
            flags = flags.SetBit(Outlined, 4);
            flags = flags.SetBit(Shadowed, 5);
            flags = flags.SetBit(Condensed, 6);
            flags = flags.SetBit(Extended, 7);
            return string.Concat(Height, Name, Color, UnderlineStyle, Family, Escapement, flags);
        }

        #endregion
    }
}
