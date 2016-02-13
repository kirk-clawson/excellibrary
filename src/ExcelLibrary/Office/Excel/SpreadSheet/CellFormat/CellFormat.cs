using ExcelLibrary.BinaryFileFormat;

namespace ExcelLibrary.SpreadSheet
{
    public class CellFormat
    {
        public CellFormat()
        {
            FormatString = FormatString.General;
            Border = new CellBorder();
            Pattern = new CellPattern();
        }

        public FormatString FormatString { get; set; }
        public CellBorder Border { get; set; }
        public CellPattern Pattern { get; set; }

        public RichTextFormat RichTextFormat { get; set; }

        public void SetBackgroundColor(ExcelColor color)
        {
            Pattern.Style = PatternStyle.Solid;
            Pattern.ForegroundColor = color;
            Pattern.BackgroundColor = ExcelColor.WindowTextForPattern;
        }

        public void ClearBackgroundColor()
        {
            Pattern.Reset();
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
            return string.Concat(FormatString, Border, Pattern);
        }

        #endregion

        internal XF CreateExtendedFormatRecord()
        {
            var xf = new XF
            {
                Attributes = 252,
                FillPattern = (byte) Pattern.Style,
                PatternColorIndex = Pattern.ForegroundColor.Index,
                PatternBackgroundColorIndex = Pattern.BackgroundColor.Index,
                TopLineStyle = (byte) Border.TopStyle,
                BottomLineStyle = (byte) Border.BottomStyle,
                LeftLineStyle = (byte) Border.LeftStyle,
                RightLineStyle = (byte) Border.RightStyle,
                TopLineColorIndex = Border.TopColor.Index,
                BottomLineColorIndex = Border.BottomColor.Index,
                LeftLineColorIndex = Border.LeftColor.Index,
                RightLineColorIndex = Border.RightColor.Index
            };

            return xf;
        }
    }
}
