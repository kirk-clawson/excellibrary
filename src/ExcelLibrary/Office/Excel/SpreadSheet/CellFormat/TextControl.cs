namespace ExcelLibrary.SpreadSheet
{
    public enum HorizontalAlignStyle : byte
    {
        General = 0,
        Left = 1,
        Centered = 2,
        Right = 3,
        Filled = 4,
        Justified = 5,
        CenteredAcrossSelection = 6,
        Distributed = 7
    }

    public enum VerticalAlignStyle : byte
    {
        Top = 0,
        Centered = 1,
        Bottom = 2,
        Justified = 3,
        Distributed = 4
    }

    public enum TextDirection : byte
    {
        Context = 0,
        LeftToRight = 1,
        RightToLeft = 2
    }

    public class TextControl
    {
        private byte _textRotation;
        private byte _indentLevel;

        public TextControl()
        {
            Reset();
        }

        public void Reset()
        {
            HorizontalAlignment = HorizontalAlignStyle.General;
            VerticalAlignment = VerticalAlignStyle.Bottom;
            TextDirection = TextDirection.Context;
            JustifyDistributed = false;
            WrapText = false;
            ShrinkToFit = false;
            IndentLevel = 0;
            RotationStyle = RotationStyle.NoRotaion;
        }

        public HorizontalAlignStyle HorizontalAlignment { get; set; }
        public VerticalAlignStyle VerticalAlignment { get; set; }
        public TextDirection TextDirection { get; set; }
        public bool JustifyDistributed { get; set; }
        public bool WrapText { get; set; }
        public bool ShrinkToFit { get; set; }

        public byte IndentLevel
        {
            get { return _indentLevel; }
            set
            {
                if (value > 15) value = 15;
                _indentLevel = value;
            }
        }

        public RotationStyle RotationStyle { get; set; }
        public byte TextRotation
        {
            get
            {
                if (RotationStyle == RotationStyle.NoRotaion || RotationStyle == RotationStyle.Stacked) return 0;
                return _textRotation;
            }
            set
            {
                if (value < 1) value = 1;
                if (value > 90) value = 90;
                _textRotation = value;
            }
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
            return string.Concat(HorizontalAlignment, VerticalAlignment, TextDirection, RotationStyle,
                                 JustifyDistributed, WrapText, ShrinkToFit, IndentLevel, TextRotation);
        }

        #endregion
    }
}