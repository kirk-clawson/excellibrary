using System.Drawing;

namespace ExcelLibrary.SpreadSheet
{
    public class ExcelColor
    {
        private ushort _index;
        private readonly Color _value;
        private readonly bool _isCustomColor;

        public ExcelColor() : this(Color.White, 0x01)
        {
        }

        public ExcelColor(Color value)
        {
            _value = value;
            _isCustomColor = true;
        }

        private ExcelColor(Color value, ushort index)
        {
            _value = value;
            _index = index;
            _isCustomColor = false;
        }

        public ushort Index
        {
            get { return _index; }
        }


        public Color Value
        {
            get { return _value; }
        }

        public bool IsCustomColor
        {
            get { return _isCustomColor; }
        }

        public void SetIndex(ushort value)
        {
            _index = value;
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
            return Value.ToString();
        }

        #endregion

        // Excel Default Palette
        public static readonly ExcelColor Black = new ExcelColor(Color.Black, 0x08);
        public static readonly ExcelColor White = new ExcelColor(Color.White, 0x09);
        public static readonly ExcelColor Red = new ExcelColor(Color.Red, 0x0A);
        public static readonly ExcelColor Green = new ExcelColor(Color.FromArgb(0x00, 0xFF, 0x00), 0x0B);
        public static readonly ExcelColor Blue = new ExcelColor(Color.Blue, 0x0C);
        public static readonly ExcelColor Yellow = new ExcelColor(Color.Yellow, 0x0D);
        public static readonly ExcelColor Magenta = new ExcelColor(Color.Magenta, 0x0E);
        public static readonly ExcelColor Cyan = new ExcelColor(Color.Cyan, 0x0F);
        public static readonly ExcelColor DarkRed = new ExcelColor(Color.FromArgb(0x80, 0x00, 0x00), 0x010);
        public static readonly ExcelColor DarkGreen = new ExcelColor(Color.Green, 0x11);
        public static readonly ExcelColor DarkBlue = new ExcelColor(Color.FromArgb(0x00, 0x00, 0x80), 0x12);
        public static readonly ExcelColor Olive = new ExcelColor(Color.Olive, 0x13);
        public static readonly ExcelColor Purple = new ExcelColor(Color.Purple, 0x14);
        public static readonly ExcelColor Teal = new ExcelColor(Color.Teal, 0x15);
        public static readonly ExcelColor Silver = new ExcelColor(Color.Silver, 0x16);
        public static readonly ExcelColor Gray = new ExcelColor(Color.Gray, 0x17);

        // System Colors
        public static readonly ExcelColor WindowTextForPattern = new ExcelColor(SystemColors.WindowText, 0x40);
        public static readonly ExcelColor WindowBackgroundForPattern = new ExcelColor(SystemColors.Window, 0x41);
        public static readonly ExcelColor SystemFace = new ExcelColor(SystemColors.Control, 0x43);
        public static readonly ExcelColor WindowTextForCharts = new ExcelColor(SystemColors.WindowText, 0x4D);
        public static readonly ExcelColor WindowBackgroundForCharts = new ExcelColor(SystemColors.Window, 0x4E);
        public static readonly ExcelColor ChartBorders = new ExcelColor(Color.Black, 0x4F);
        public static readonly ExcelColor ToolTipBackground = new ExcelColor(SystemColors.Info, 0x50);
        public static readonly ExcelColor ToolTipText = new ExcelColor(SystemColors.InfoText, 0x51);
        public static readonly ExcelColor WindowTextForFonts = new ExcelColor(SystemColors.WindowText, 0x7FFF);
    }
}