using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelLibrary.SpreadSheet;

namespace ExcelLibrary.SampleUsage
{
    class Program
    {
        static void Main(string[] args)
        {
            var underTest = new Workbook();
            var sheet = new Worksheet("Sheet 1");

            var styleA = new CellFormat();
            styleA.SetBackgroundColor(ExcelColor.Red);
            styleA.Border.DiagonalUp = true;
            styleA.Border.DiagonalDown = true;
            styleA.Border.DiagonalStyle = CellBorderStyle.Thin;
            var styleB = new CellFormat
            {
                Pattern =
                {
                    Style = PatternStyle.LightDown,
                    ForegroundColor = ExcelColor.Blue
                }
            };
            var styleC = new CellFormat();
            styleC.SetBackgroundColor(ExcelColor.Silver);
            styleC.Border = CellBorder.MediumBox;
            styleC.Font.Bold = true;

            for (var i = 0; i < 100; i++)
            {
                var cellA = new Cell("Abcde");
                var cellB = new Cell(1234);
                var cellC = new Cell(string.Format("This is row {0:000}", i));
                if (i%2 == 0)
                {
                    cellA.CellFormat = styleA;
                    cellB.CellFormat = styleB;
                    cellC.CellFormat = styleC;
                }
                else
                {
                    cellB.CellFormat.TextControl.RotationStyle = RotationStyle.CounterClockwise;
                    cellB.CellFormat.TextControl.TextRotation = 45;
                    cellC.VerticalAlignment = VerticalAlignStyle.Centered;
                    cellC.CellFormat.Font.Name = "Times New Roman";
                    cellC.CellFormat.Font.Family = FontFamilyType.Roman;
                    cellC.CellFormat.Font.Height = 240;
                }
                sheet.Cells[i, 0] = cellA;
                sheet.Cells[i, 1] = cellB;
                sheet.Cells[i, 2] = cellC;
                sheet.Cells.ColumnWidth[2] = 256 * 15;
            }
            underTest.Worksheets.Add(sheet);
            var filename = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "test.xls");
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }
            underTest.Save(filename);
        }
    }
}
