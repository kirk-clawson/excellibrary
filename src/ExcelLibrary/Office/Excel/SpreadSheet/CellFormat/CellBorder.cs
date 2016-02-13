﻿namespace ExcelLibrary.SpreadSheet
{
    public enum CellBorderStyle : byte
    {
        NoLine = 0,
        Thin = 1,
        Medium = 2,
        Dashed = 3,
        Dotted = 4,
        Thick = 5,
        Double = 6,
        Hair = 7,
        MediumDashed = 8,
        ThinDashDot = 9,
        MediumDashDot = 10,
        ThinDashDotDot = 11,
        MediumDashDotDot = 12,
        SlantedDashDto = 13
    }

    public class CellBorder
    {
        public CellBorderStyle TopStyle;
        public ExcelColor TopColor;
        public CellBorderStyle BottomStyle;
        public ExcelColor BottomColor;
        public CellBorderStyle LeftStyle;
        public ExcelColor LeftColor;
        public CellBorderStyle RightStyle;
        public ExcelColor RightColor;

        public CellBorder()
        {
            Reset();
        }

        public CellBorder(CellBorderStyle lineStyle) : this()
        {
            TopStyle = lineStyle;
            BottomStyle = lineStyle;
            LeftStyle = lineStyle;
            RightStyle = lineStyle;
        }

        public void Reset()
        {
            TopStyle = CellBorderStyle.NoLine;
            BottomStyle = CellBorderStyle.NoLine;
            LeftStyle = CellBorderStyle.NoLine;
            RightStyle = CellBorderStyle.NoLine;
            TopColor = ExcelColor.WindowTextForPattern;
            BottomColor = ExcelColor.WindowTextForPattern;
            LeftColor = ExcelColor.WindowTextForPattern;
            RightColor = ExcelColor.WindowTextForPattern;
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
            return string.Concat(TopStyle, BottomStyle, LeftStyle, RightStyle, TopColor, BottomColor, LeftColor, RightColor);
        }

        #endregion

        public static readonly CellBorder MediumBox = new CellBorder(CellBorderStyle.Medium);
        public static readonly CellBorder MediumBottom = new CellBorder { BottomStyle = CellBorderStyle.Medium };
        public static readonly CellBorder MediumTopBottom = new CellBorder { BottomStyle = CellBorderStyle.Medium, TopStyle = CellBorderStyle.Medium };
        public static readonly CellBorder ThinBox = new CellBorder(CellBorderStyle.Thin);
    }
}
