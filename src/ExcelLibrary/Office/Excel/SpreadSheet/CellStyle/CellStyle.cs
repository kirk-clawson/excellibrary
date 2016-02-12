using System.Drawing;
using ExcelLibrary.BinaryFileFormat;

namespace ExcelLibrary.SpreadSheet
{
    public class CellStyle
    {
        public Color BackgroundColor = Color.White;
        public CellBorder Border;
        public RichTextFormat RichTextFormat;
    }
}
