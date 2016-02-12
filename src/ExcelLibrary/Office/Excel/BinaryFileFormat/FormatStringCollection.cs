using System;
using System.Collections.Generic;
using System.Linq;
using ExcelLibrary.SpreadSheet;

namespace ExcelLibrary.BinaryFileFormat
{
    /// <summary>
    /// Excel built-in cell format 
    /// source:             
    ///  http://svn.apache.org/viewvc/poi/trunk/src/java/org/apache/poi/hssf/usermodel/HSSFDataFormat.java?view=markup  
    ///  http://sc.openoffice.org/excelfileformat.pdf 
    /// </summary>
    public class FormatStringCollection
    {
        private readonly Dictionary<ushort, FormatString> _lookupTable;

        public FormatStringCollection()
        {
            _lookupTable = new Dictionary<ushort, FormatString>
            {
                {0, new FormatString(FormatStringType.General, "General")},
                {1, new FormatString(FormatStringType.Number, "0")},
                {2, new FormatString(FormatStringType.Number, "0.00")},
                {3, new FormatString(FormatStringType.Number, "#,##0")},
                {4, new FormatString(FormatStringType.Number, "#,##0.00")},
                {5, new FormatString(FormatStringType.Currency, "($#,##0_);($#,##0)")},
                {6, new FormatString(FormatStringType.Currency, "($#,##0_);[Red]($#,##0)")},
                {7, new FormatString(FormatStringType.Currency, "($#,##0.00);($#,##0.00)")},
                {8, new FormatString(FormatStringType.Currency, "($#,##0.00_);[Red]($#,##0.00)")},
                {9, new FormatString(FormatStringType.Percentage, "0%")},
                {10, new FormatString(FormatStringType.Percentage, "0.00%")},
                {11, new FormatString(FormatStringType.Scientific, "0.00E+00")},
                {12, new FormatString(FormatStringType.Fraction, "# ?/?")},
                {13, new FormatString(FormatStringType.Fraction, "# ??/??")},
                {14, new FormatString(FormatStringType.Date, "m/d/yy")},
                {15, new FormatString(FormatStringType.Date, "d-mmm-yy")},
                {16, new FormatString(FormatStringType.Date, "d-mmm")},
                {17, new FormatString(FormatStringType.Date, "mmm-yy")},
                {18, new FormatString(FormatStringType.Time, "h:mm AM/PM")},
                {19, new FormatString(FormatStringType.Time, "h:mm:ss AM/PM")},
                {20, new FormatString(FormatStringType.Time, "h:mm")},
                {21, new FormatString(FormatStringType.Time, "h:mm:ss")},
                {22, new FormatString(FormatStringType.DateTime, "m/d/yy h:mm")},
                {37, new FormatString(FormatStringType.Accounting, "(#,##0_);(#,##0)")},
                {38, new FormatString(FormatStringType.Accounting, "(#,##0_);[Red](#,##0)")},
                {39, new FormatString(FormatStringType.Accounting, "(#,##0.00_);(#,##0.00)")},
                {40, new FormatString(FormatStringType.Accounting, "(#,##0.00_);[Red](#,##0.00)")},
                {41, new FormatString(FormatStringType.Currency, "_(*#,##0_);_(*(#,##0);_(* \"-\"_);_(@_)")},
                {42, new FormatString(FormatStringType.Currency, "_($*#,##0_);_($*(#,##0);_($* \"-\"_);_(@_)")},
                {43, new FormatString(FormatStringType.Currency, "_(*#,##0.00_);_(*(#,##0.00);_(*\"-\"??_);_(@_)")},
                {44, new FormatString(FormatStringType.Currency, "_($*#,##0.00_);_($*(#,##0.00);_($*\"-\"??_);_(@_)")},
                {45, new FormatString(FormatStringType.Time, "mm:ss")},
                {46, new FormatString(FormatStringType.Time, "[h]:mm:ss")},
                {47, new FormatString(FormatStringType.Time, "mm:ss.0")},
                {48, new FormatString(FormatStringType.Scientific, "##0.0E+0")},
                {49, new FormatString(FormatStringType.Text, "@")}
            };
        }

        public void Add(FORMAT record)
        {
            if (record == null) return;
            // Built-in cell formula may change due to regional settings            
            // therefore, we allow caller to replace built-in cell format      
            if (_lookupTable.ContainsKey(record.FormatIndex))
            {
                FormatString oldCellFormat = _lookupTable[record.FormatIndex];
                _lookupTable[record.FormatIndex] = new FormatString(oldCellFormat.FormatType, record.FormatString);
            }
            else
            {
                _lookupTable.Add(record.FormatIndex, new FormatString(FormatStringType.Custom, record.FormatString));
            }
        }

        public FormatString this[ushort formatIndex]
        {
            get
            {
                if (_lookupTable.ContainsKey(formatIndex))
                {
                    return _lookupTable[formatIndex];
                }
                throw new KeyNotFoundException("Unable to find specific cell format");
            }
        }

        public UInt16 GetFormatIndex(string formatString)
        {
            foreach (KeyValuePair<ushort, FormatString> cellFormat in _lookupTable.Where(kv => formatString == kv.Value.Value))
            {
                return cellFormat.Key;
            }
            return UInt16.MaxValue;
        }
    }
}



