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
        public SST SharedStringTable;

        public DateTime BaseDate;

        public ColorPalette ColorPalette = new ColorPalette();

        public List<FORMAT> FormatRecords = new List<FORMAT>();

        public List<XF> ExtendedFormats = new List<XF>();

        public CellFormatCollection CellFormats = new CellFormatCollection();

        public FastSearchList<Image> Images = new FastSearchList<Image>();

        public List<FONT> Fonts = new List<FONT>();

        public SharedResource()
        {
        }

        public SharedResource(bool newbook)
        {
            FONT font = new FONT();
            font.Height = 200;
            font.OptionFlags = 0;
            font.ColorIndex = 32767;
            font.Weight = 400;
            font.Escapement = 0;
            font.Underline = 0;
            font.CharacterSet = 1;
            font.Name = "Arial";
            //Fonts.Add(font);

            for (ushort i = 0; i < 21; i++) // required by MS Excel 2003
            {
                XF xf = new XF();
                xf.Attributes = 252;
                xf.CellProtection = 65524;
                xf.PatternColorIndex = 64;
                xf.PatternBackgroundColorIndex = 130;
                xf.FontIndex = 0;
                xf.FormatIndex = i;
                ExtendedFormats.Add(xf);
            }

            _maxNumberFormatIndex = 163;
            GetXFIndex(CellFormat.General, new CellStyle());

            SharedStringTable = new SST();
        }

        public string GetStringFromSST(int index)
        {
            if (SharedStringTable != null)
            {
                return SharedStringTable.StringList[index];
            }
            return null;
        }

        public int GetSSTIndex(string text)
        {
            SharedStringTable.TotalOccurance++;
            int index = SharedStringTable.StringList.IndexOf(text);
            if (index == -1)
            {
                SharedStringTable.StringList.Add(text);
                return SharedStringTable.StringList.Count - 1;
            }
            else
            {
                return index;
            }
        }

        public double EncodeDateTime(DateTime value)
        {
            double days = (value - BaseDate).TotalDays;
            return days;
        }

        private readonly Dictionary<string, int> _numberFormatXfIndice = new Dictionary<string, int>();
        private ushort _maxNumberFormatIndex;
        internal int GetXFIndex(CellFormat cellFormat, CellStyle cellStyle)
        {
            string formatKey = GetXfKey(cellFormat, cellStyle);
            if (_numberFormatXfIndice.ContainsKey(formatKey))
            {
                return _numberFormatXfIndice[formatKey];
            }
            else
            {
                UInt16 formatIndex = CellFormats.GetFormatIndex(cellFormat.FormatString);
                if (formatIndex == UInt16.MaxValue)
                {
                    formatIndex = _maxNumberFormatIndex++;
                }

                FORMAT format = new FORMAT();
                format.FormatIndex = formatIndex;
                format.FormatString = cellFormat.FormatString;
                FormatRecords.Add(format);

                XF xf = new XF();
                xf.Attributes = 252;
                xf.CellProtection = 0;
                xf.PatternColorIndex = 64;
                xf.PatternBackgroundColorIndex = 65;
                xf.FontIndex = 0;
                xf.FormatIndex = formatIndex;
                if (cellStyle != null && cellStyle.BackgroundColor != Color.White)
                {
                    // 1 is the solid fill pattern value
                    xf.FillPattern = 1;
                    xf.PatternColorIndex = ColorPalette.BuiltInIndexes[cellStyle.BackgroundColor];
                }
                ExtendedFormats.Add(xf);

                int numberFormatXFIndex = ExtendedFormats.Count - 1;
                _numberFormatXfIndice.Add(formatKey, numberFormatXFIndex);

                return numberFormatXFIndex;
            }
        }

        private string GetXfKey(CellFormat format, CellStyle style)
        {
            string formatString = format.FormatString;
            string styleString = style?.BackgroundColor.ToString() ?? "";
            return formatString + styleString;
        }
    }
}
