using System.Drawing;

namespace ExcelLibrary.SpreadSheet
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
        public Color TopColor;
        public CellBorderStyle BottomStyle;
        public Color BottomColor;
        public CellBorderStyle LeftStyle;
        public Color LeftColor;
        public CellBorderStyle RightStyle;
        public Color RightColor;

        public CellBorder() : this(CellBorderStyle.NoLine)
        {
        }

        public CellBorder(CellBorderStyle lineStyle)
        {
            TopStyle = lineStyle;
            BottomStyle = lineStyle;
            LeftStyle = lineStyle;
            RightStyle = lineStyle;
            TopColor = Color.Black;
            BottomColor = Color.Black;
            LeftColor = Color.Black;
            RightColor = Color.Black;
        }

        public static readonly CellBorder MediumBox = new CellBorder(CellBorderStyle.Medium);
        public static readonly CellBorder MediumBottom = new CellBorder { BottomStyle = CellBorderStyle.Medium };
        public static readonly CellBorder MediumTopBottom = new CellBorder { BottomStyle = CellBorderStyle.Medium, TopStyle = CellBorderStyle.Medium };
        public static readonly CellBorder ThinBox = new CellBorder(CellBorderStyle.Thin);

    }
}
