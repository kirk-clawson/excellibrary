namespace ExcelLibrary.SpreadSheet
{
    public class FormatString
    {
        private FormatString() : this (FormatStringType.General, "General")
        {
        }

        public FormatString(FormatStringType type, string fmt)
        {
            FormatType = type;
            Value = fmt;
        }

        public FormatStringType FormatType { get; set; }
        public string Value { get; set; }

        #region Overrides of Object

        /// <summary>
        /// Returns a string that represents the current object.
        /// </summary>
        /// <returns>
        /// A string that represents the current object.
        /// </returns>
        public override string ToString()
        {
            return Value;
        }

        #endregion

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
