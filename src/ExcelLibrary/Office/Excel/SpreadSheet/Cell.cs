using System;
using ExcelLibrary.BinaryFileFormat;

namespace ExcelLibrary.SpreadSheet
{
    public class Cell
    {
        private object _value;

        internal SharedResource SharedResource;

        public static readonly Cell EmptyCell = new Cell(null);

        public Cell(object value)
        {
            _value = value;
            CellFormat = new CellFormat();
        }

        public Cell(object value, string formatString)
        {
            _value = value;
            CellFormat = new CellFormat(); //TODO: get format string into this structure
        }

        public Cell(object value, CellFormat cellFormat)
        {
            _value = value;
            CellFormat = cellFormat;
        }

        public override string ToString()
        {
            return StringValue;
        }

        public bool IsEmpty
        {
            get { return this == EmptyCell; }
        }

        public object Value
        {
            get
            {
                return _value;
            }
            set
            {
                if (IsEmpty) throw new Exception("Can not set value to an empty cell.");
                _value = value;
            }
        }

        public string StringValue
        {
            get
            {
                return _value == null ? string.Empty : _value.ToString();
            }
        }

        public DateTime DateTimeValue
        {
            get
            {
                if (_value is double)
                {
                    var days = (double)_value;
                    //Excel counts an extra day for 1900-Feb-29. In reality, 1900 is not a leap year.
                    if (SharedResource.BaseDate == DateTime.Parse("1899-12-31") && days > 59)
                    {
                        days--;
                    }
                    return SharedResource.BaseDate.AddDays(days);
                }
                var s = _value as string;
                if (s != null)
                {
                    return DateTime.Parse(s);
                }
                if (_value is DateTime)
                {
                    return (DateTime)_value;
                }
                throw new Exception("Invalid DateTime Cell.");
            }
            set
            {
                _value = value;
            }
        }

        public string FormatString
        {
            get { return CellFormat.Format.Value; }
            set { CellFormat.Format.Value = value; }
        }

        public CellFormat CellFormat { get; set; }

        public FONT GetFontForCharacter(UInt16 charIndex)
        {
            return WorksheetDecoder.getFontForCharacter(this, charIndex);
        }
    }
}
