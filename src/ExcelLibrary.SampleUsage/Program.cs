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

            var styleA = new CellStyle {BackgroundColor = Color.Red};
            var styleB = new CellStyle {BackgroundColor = Color.Green};
            var styleC = new CellStyle {BackgroundColor = Color.Blue};

            for (int i = 0; i < 100; i++)
            {
                var cellA = new Cell("Abcd");
                var cellB = new Cell(1234);
                var cellC = new Cell(DateTime.Now);
                if (i%2 == 0)
                {
                    cellA.Style = styleA;
                    cellB.Style = styleB;
                    cellC.Style = styleC;
                }
                sheet.Cells[i, 0] = cellA;
                sheet.Cells[i, 1] = cellB;
                sheet.Cells[i, 2] = cellC;
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
