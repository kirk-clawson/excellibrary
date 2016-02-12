using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using QiHe.CodeLib;
using ExcelLibrary.SpreadSheet;
using Image = ExcelLibrary.SpreadSheet.Image;

namespace ExcelLibrary.BinaryFileFormat
{
    public class SharedResource
    {
        private readonly Dictionary<string, int> _numberFormatXfIndice = new Dictionary<string, int>();
        private ushort _maxNumberFormatIndex;

        public SST SharedStringTable;
        public DateTime BaseDate;
        public ColorPalette ColorPalette = new ColorPalette();
        public List<FORMAT> FormatRecords = new List<FORMAT>();
        public List<XF> ExtendedFormats = new List<XF>();
        public FormatStringCollection FormatStrings = new FormatStringCollection();
        public FastSearchList<Image> Images = new FastSearchList<Image>();
        public List<FONT> Fonts = new List<FONT>();
        
        public SharedResource()
        {
        }

        public SharedResource(bool newbook)
        {
            for (ushort i = 0; i < 21; i++) // required by MS Excel 2003
            {
                var xf = new XF
                {
                    Attributes = 252,
                    CellProtection = 65524,
                    PatternColorIndex = 64,
                    PatternBackgroundColorIndex = 130,
                    FontIndex = 0,
                    FormatIndex = i
                };
                ExtendedFormats.Add(xf);
            }

            _maxNumberFormatIndex = 163;
            GetXFIndex(new CellFormat());

            SharedStringTable = new SST();
        }

        public string GetStringFromSST(int index)
        {
            return SharedStringTable != null ? SharedStringTable.StringList[index] : null;
        }

        public int GetSSTIndex(string text)
        {
            SharedStringTable.TotalOccurance++;
            int index = SharedStringTable.StringList.IndexOf(text);
            if (index != -1) return index;

            SharedStringTable.StringList.Add(text);
            return SharedStringTable.StringList.Count - 1;
        }

        public double EncodeDateTime(DateTime value)
        {
            double days = (value - BaseDate).TotalDays;
            return days;
        }

        
        internal int GetXFIndex(CellFormat cellFormat)
        {
            string formatKey = cellFormat.ToString();
            if (_numberFormatXfIndice.ContainsKey(formatKey))
            {
                return _numberFormatXfIndice[formatKey];
            }
            else
            {
                UInt16 formatIndex = FormatStrings.GetFormatIndex(cellFormat.Format.Value);
                if (formatIndex == UInt16.MaxValue)
                {
                    formatIndex = _maxNumberFormatIndex++;
                }

                FORMAT format = new FORMAT();
                format.FormatIndex = formatIndex;
                format.FormatString = cellFormat.Format.Value;
                FormatRecords.Add(format);

                XF xf = new XF();
                xf.Attributes = 252;
                xf.CellProtection = 0;
                xf.PatternColorIndex = 64;
                xf.PatternBackgroundColorIndex = 65;
                xf.FontIndex = 0;
                xf.FormatIndex = formatIndex;
                if (cellFormat.BackgroundColor != Color.White)
                {
                    // 1 is the solid fill pattern value
                    xf.FillPattern = 1;
                    xf.PatternColorIndex = ColorPalette.BuiltInIndexes[cellFormat.BackgroundColor];
                }
                if (cellFormat.Border != null)
                {
                    xf.TopLineStyle = (byte)cellFormat.Border.TopStyle;
                    xf.BottomLineStyle = (byte)cellFormat.Border.BottomStyle;
                    xf.LeftLineStyle = (byte)cellFormat.Border.LeftStyle;
                    xf.RightLineStyle = (byte)cellFormat.Border.RightStyle;
                    xf.TopLineColorIndex = ColorPalette.BuiltInIndexes[cellFormat.Border.TopColor];
                    xf.BottomLineColorIndex = ColorPalette.BuiltInIndexes[cellFormat.Border.BottomColor];
                    xf.LeftLineColorIndex = ColorPalette.BuiltInIndexes[cellFormat.Border.LeftColor];
                    xf.RightLineColorIndex = ColorPalette.BuiltInIndexes[cellFormat.Border.RightColor];
                }
                ExtendedFormats.Add(xf);

                _numberFormatXfIndice.Add(formatKey, ExtendedFormats.Count - 1);

                return ExtendedFormats.Count - 1;
            }
        }
    }
}
