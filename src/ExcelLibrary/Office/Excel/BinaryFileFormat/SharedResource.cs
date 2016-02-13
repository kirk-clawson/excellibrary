using System;
using System.Collections.Generic;
using QiHe.CodeLib;
using ExcelLibrary.SpreadSheet;
using Image = ExcelLibrary.SpreadSheet.Image;

namespace ExcelLibrary.BinaryFileFormat
{
    public class SharedResource
    {
        private readonly Dictionary<string, int> _cellFormatIndexLookup = new Dictionary<string, int>();
        private ushort _maxNumberFormatIndex;

        public SST SharedStringTable;
        public DateTime BaseDate;
        public List<FORMAT> StringFormatRecords = new List<FORMAT>();
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
            if (_cellFormatIndexLookup.ContainsKey(formatKey))
            {
                return _cellFormatIndexLookup[formatKey];
            }

            UInt16 formatIndex = FormatStrings.GetFormatIndex(cellFormat.FormatString.Value);
            if (formatIndex == UInt16.MaxValue)
            {
                formatIndex = _maxNumberFormatIndex++;
            }

            var format = new FORMAT
            {
                FormatIndex = formatIndex,
                FormatString = cellFormat.FormatString.Value
            };
            StringFormatRecords.Add(format);

            XF xf = cellFormat.CreateExtendedFormatRecord();
            xf.FontIndex = 0;
            xf.FormatIndex = formatIndex;
            ExtendedFormats.Add(xf);

            _cellFormatIndexLookup.Add(formatKey, ExtendedFormats.Count - 1);

            return ExtendedFormats.Count - 1;
        }
    }
}
