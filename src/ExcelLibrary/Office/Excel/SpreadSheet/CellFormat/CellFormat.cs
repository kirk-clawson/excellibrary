using ExcelLibrary.BinaryFileFormat;

namespace ExcelLibrary.SpreadSheet
{
    public class CellFormat
    {
        public CellFormat()
        {
            Font = new CellFont();
            FormatString = new FormatString();
            Border = new CellBorder();
            Pattern = new CellPattern();
            TextControl = new TextControl();
        }

        public bool LockCell { get; set; }
        public bool HideFormula { get; set; }

        public CellFont Font { get; set; }
        public FormatString FormatString { get; set; }
        public CellBorder Border { get; set; }
        public CellPattern Pattern { get; set; }
        public TextControl TextControl { get; set; }

        public RichTextFormat RichTextFormat { get; set; }

        public void SetBackgroundColor(ExcelColor color)
        {
            Pattern.Style = PatternStyle.Solid;
            Pattern.ForegroundColor = color;
            Pattern.BackgroundColor = ExcelColor.WindowTextForPattern;
        }

        public ExcelColor GetBackgroundColor()
        {
            return Pattern.Style == PatternStyle.Solid ? Pattern.ForegroundColor : null;
        }

        public void ClearBackgroundColor()
        {
            Pattern.Reset();
        }

        public void Reset()
        {
            FormatString.Reset();
            Border.Reset();
            Pattern.Reset();
            TextControl.Reset();
            LockCell = false;
            HideFormula = false;
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
            return string.Concat(Font, FormatString, Border, Pattern, TextControl, LockCell, HideFormula);
        }

        #endregion
    }
}
