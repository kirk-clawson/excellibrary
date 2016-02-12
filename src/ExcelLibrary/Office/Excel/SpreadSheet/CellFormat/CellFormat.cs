using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using ExcelLibrary.BinaryFileFormat;

namespace ExcelLibrary.SpreadSheet
{
    public class CellFormat
    {
        public CellFormat()
        {
            Format = FormatString.General;
        }

        public FormatString Format { get; set; }
        public Color BackgroundColor = Color.White;
        public CellBorder Border;
        public RichTextFormat RichTextFormat;

        public override string ToString()
        {
            return Format.Value + BackgroundColor.ToString();
        }
    }
}
