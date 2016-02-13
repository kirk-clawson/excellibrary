namespace ExcelLibrary.SpreadSheet
{
    public enum PatternStyle : byte
    {
        NoPattern = 0x00,
        Solid = 0x01,
        Gray50 = 0x02,
        Gray75 = 0x03,
        Gray25 = 0x04,
        Horizontal = 0x05,
        Vertical = 0x06,
        Down = 0x07,
        Up = 0x08,
        CheckerBoard = 0x09,
        SemiGray75 = 0x0A,
        LightHorizontal = 0x0B,
        LightVertical = 0x0C,
        LightDown = 0x0D,
        LightUp = 0x0E,
        Grid = 0x0F,
        CrissCross = 0x10,
        Gray16 = 0x11,
        Gray8 = 0x12
    }

    public class CellPattern
    {
        public PatternStyle Style;
        public ExcelColor ForegroundColor;
        public ExcelColor BackgroundColor;

        public CellPattern()
        {
            Reset();
        }

        public void Reset()
        {
            Style = PatternStyle.NoPattern;
            ForegroundColor = ExcelColor.WindowTextForPattern;
            BackgroundColor = ExcelColor.WindowBackgroundForPattern;
        }

        #region Overrides of Object

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns>
        /// A string that represents the current object.
        /// </returns>
        public override string ToString()
        {
            return string.Concat(Style, ForegroundColor, BackgroundColor);
        }

        #endregion

        public static readonly CellPattern Default = new CellPattern();
    }
}
