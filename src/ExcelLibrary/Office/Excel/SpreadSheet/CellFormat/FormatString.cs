namespace ExcelLibrary.SpreadSheet
{
    public class FormatString
    {

        private FormatStringType _formatType = FormatStringType.General;
        private string _value = "General";

        public FormatString()
        {
            DataChanged = false;
        }

        public FormatString(FormatStringType type, string fmt)
        {
            FormatType = type;
            Value = fmt;
        }

        public FormatStringType FormatType
        {
            get { return _formatType; }
            set
            {
                if (value != _formatType)
                {
                    DataChanged = true;
                }
                _formatType = value;
            }
        }

        public string Value
        {
            get { return _value; }
            set
            {
                if (value != _value)
                {
                    DataChanged = true;
                }
                _value = value;
            }
        }

        public bool DataChanged { get; set; }

        /// <summary>
        /// Default Format.
        /// </summary>
        public static readonly FormatString General = new FormatString();
        
        /// <summary>
        /// Format the DateTime with: "YYYY-MM-DD" e.g: 2009-02-18
        /// </summary>
        public static readonly FormatString Date = new FormatString(FormatStringType.Date, @"YYYY\-MM\-DD");
        
        /// <summary>
        /// Format the DateTime with: "HH:mm:ss" e.g: 14:45:00
        /// </summary>
        /// <example>Format the DateTime with: "HH:mm:ss" e.g: 14:45:00</example>
        public static readonly FormatString Time = new FormatString(FormatStringType.Time, "HH:mm:ss");
        
        /// <summary>
        /// Format the number with: "#,###.00000" e.g: 100.12345
        /// </summary>
        /// <example>Format the number with: "#,###.00000" e.g: 100.12345</example>
        public static readonly FormatString Engineer = new FormatString(FormatStringType.Scientific, "#,###.00000");
    }
}
