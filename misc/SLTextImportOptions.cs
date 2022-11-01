using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SpreadsheetLight
{
    /// <summary>
    /// The type of data fields to be imported, whether by delimiters/separators or in fixed width.
    /// </summary>
    public enum SLTextImportDataFieldTypeValues
    {
        /// <summary>
        /// Data is separated by character delimiters.
        /// </summary>
        Delimited = 0,
        /// <summary>
        /// Data is separated by fixed width columns.
        /// </summary>
        FixedWidth
    }

    /// <summary>
    /// The type of column data format.
    /// </summary>
    public enum SLTextImportColumnFormatValues
    {
        /// <summary>
        /// Numeric values will be converted to numbers, date values to dates and remaining values to text.
        /// </summary>
        General = 0,
        /// <summary>
        /// Text format.
        /// </summary>
        Text,
        /// <summary>
        /// The value will be parsed as a date in the order of month, day, year.
        /// Failing that, any given custom date formats will be used to parse the value.
        /// Failing that, the value is parse generically as a date.
        /// And failing that, the value is set as text.
        /// </summary>
        DateMDY,
        /// <summary>
        /// The value will be parsed as a date in the order of day, month, year.
        /// Failing that, any given custom date formats will be used to parse the value.
        /// Failing that, the value is parse generically as a date.
        /// And failing that, the value is set as text.
        /// </summary>
        DateDMY,
        /// <summary>
        /// The value will be parsed as a date in the order of year, month, day.
        /// Failing that, any given custom date formats will be used to parse the value.
        /// Failing that, the value is parse generically as a date.
        /// And failing that, the value is set as text.
        /// </summary>
        DateYMD,
        /// <summary>
        /// The value will be parsed as a date in the order of month, year, day.
        /// Failing that, any given custom date formats will be used to parse the value.
        /// Failing that, the value is parse generically as a date.
        /// And failing that, the value is set as text.
        /// </summary>
        DateMYD,
        /// <summary>
        /// The value will be parsed as a date in the order of day, year, month.
        /// Failing that, any given custom date formats will be used to parse the value.
        /// Failing that, the value is parse generically as a date.
        /// And failing that, the value is set as text.
        /// </summary>
        DateDYM,
        /// <summary>
        /// The value will be parsed as a date in the order of year, day, month.
        /// Failing that, any given custom date formats will be used to parse the value.
        /// Failing that, the value is parse generically as a date.
        /// And failing that, the value is set as text.
        /// </summary>
        DateYDM,
        /// <summary>
        /// This column will be skipped.
        /// </summary>
        Skip
    }

    /// <summary>
    /// Text import options for importing text data.
    /// </summary>
    public class SLTextImportOptions
    {
        /// <summary>
        /// Indicates if fields are separated by character delimiters or are of fixed width.
        /// The default is Delimited.
        /// </summary>
        public SLTextImportDataFieldTypeValues DataFieldType { get; set; }

        // Excel by default only has the tab delimiter turned on.
        // LibreOffice Calc by default has the tab, comma and semicolon delimiters turned on.
        // I'm gonna follow Excel. Having the comma delimiter makes sense. Because you know,
        // the C in CSV stands for "comma". But if the data contains something like
        // "1,234,567.89" then the comma is a hindrance. We want the 1.234 million.

        private int iDefaultFixedWidth;
        /// <summary>
        /// The default number of characters when columns are of fixed width.
        /// If no width is set for a column, this will be used. By default, this is 8 characters.
        /// </summary>
        public int DefaultFixedWidth
        {
            get { return iDefaultFixedWidth; }
            set
            {
                if (value >= 1) iDefaultFixedWidth = value;
            }
        }

        /// <summary>
        /// Indicates if a tab character is a delimiter. By default, this is true.
        /// </summary>
        public bool UseTabDelimiter { get; set; }

        /// <summary>
        /// Indicates if a semicolon is a delimiter. By default, this is false.
        /// </summary>
        public bool UseSemicolonDelimiter { get; set; }

        /// <summary>
        /// Indicates if a comma is a delimiter. By default, this is false.
        /// </summary>
        public bool UseCommaDelimiter { get; set; }

        /// <summary>
        /// Indicates if a space character is a delimiter. By default, this is false.
        /// </summary>
        public bool UseSpaceDelimiter { get; set; }

        /// <summary>
        /// Indicates if a custom character is used as a delimiter. By default, this is false. Use the CustomDelimiter property to set the custom delimiter character.
        /// </summary>
        public bool UseCustomDelimiter { get; set; }

        /// <summary>
        /// The custom delimiter character. This is used only when UseCustomDelimiter is true.
        /// </summary>
        public char CustomDelimiter { get; set; }

        /// <summary>
        /// Indicates if consecutive delimiters are treated as one.
        /// </summary>
        public bool MergeDelimiters { get; set; }

        /// <summary>
        /// Indicates if data enclosed within text qualifiers is taken as text.
        /// The default is true.
        /// </summary>
        public bool HasTextQualifier { get; set; }

        /// <summary>
        /// Data enclosed within this qualifier will automatically be taken as text. The text qualifier
        /// will be removed. The default is the double quote character.
        /// </summary>
        public char TextQualifier { get; set; }

        private int iImportStartRowIndex;
        /// <summary>
        /// The row in the text data source to begin importing.
        /// </summary>
        public int ImportStartRowIndex
        {
            get { return iImportStartRowIndex; }
            set
            {
                // because 0 and negative numbers don't make sense
                if (value >= 1) iImportStartRowIndex = value;
            }
        }

        /// <summary>
        /// The culture used for parsing numbers and dates. The default is the InvariantCulture.
        /// </summary>
        public CultureInfo Culture { get; set; }

        /// <summary>
        /// The number styles used for parsing numeric data. The default is NumberStyles.Any.
        /// </summary>
        public NumberStyles NumberStyles { get; set; }

        /// <summary>
        /// The encoding used to read the data source. The default is Encoding.Default.
        /// </summary>
        public Encoding Encoding { get; set; }

        /// <summary>
        /// Indicates if space characters in the data source are preserved (after data column separation by delimiters). By default, this is true.
        /// </summary>
        public bool PreserveSpace { get; set; }

        internal Dictionary<int, SLTextImportColumnFormatValues> dictColumnFormat;
        internal List<string> listCustomDateFormats;

        internal Dictionary<int, int> dictFixedWidth;

        /// <summary>
        /// Initializes an instance of SLTextImportOptions, and assuming that the data source is character delimited.
        /// </summary>
        public SLTextImportOptions()
        {
            this.SetAllNull(SLTextImportDataFieldTypeValues.Delimited);
        }

        /// <summary>
        /// Initializes an instance of SLTextImportOptions.
        /// </summary>
        /// <param name="DataFieldType">Whether the data source is character delimited or of fixed width.</param>
        public SLTextImportOptions(SLTextImportDataFieldTypeValues DataFieldType)
        {
            this.SetAllNull(DataFieldType);
        }

        private void SetAllNull(SLTextImportDataFieldTypeValues DataFieldType)
        {
            this.DataFieldType = DataFieldType;
            this.iDefaultFixedWidth = 8;
            this.UseTabDelimiter = true;
            this.UseSemicolonDelimiter = false;
            this.UseCommaDelimiter = false;
            this.UseSpaceDelimiter = false;
            this.UseCustomDelimiter = false;
            this.CustomDelimiter = ' ';
            this.MergeDelimiters = false;
            this.HasTextQualifier = true;
            this.TextQualifier = '"';
            this.iImportStartRowIndex = 1;
            this.Culture = CultureInfo.InvariantCulture;
            this.NumberStyles = System.Globalization.NumberStyles.Any;
            this.Encoding = Encoding.Default;
            this.PreserveSpace = true;
            this.dictColumnFormat = new Dictionary<int, SLTextImportColumnFormatValues>();
            this.listCustomDateFormats = new List<string>();
            this.dictFixedWidth = new Dictionary<int, int>();
        }

        /// <summary>
        /// Set the column data format type.
        /// </summary>
        /// <param name="ColumnIndex">The column index in the data source. This is 1-based indexing, so it's 1 for the 1st data source column, 2 for the 2nd data source column and so on.</param>
        /// <param name="ColumnFormat">The column data format type.</param>
        public void SetColumnFormat(int ColumnIndex, SLTextImportColumnFormatValues ColumnFormat)
        {
            if (ColumnIndex >= 1)
            {
                this.dictColumnFormat[ColumnIndex] = ColumnFormat;
            }
        }

        /// <summary>
        /// Skip a particular data source column. This is equivalent to using the SetColumnFormat() function with a Skip data format type.
        /// </summary>
        /// <param name="ColumnIndex">The data source column to skip. This is 1-based indexing, so it's 1 for the 1st data source column, 2 for the 2nd data source column and so on.</param>
        public void SkipColumn(int ColumnIndex)
        {
            if (ColumnIndex >= 1)
            {
                this.dictColumnFormat[ColumnIndex] = SLTextImportColumnFormatValues.Skip;
            }
        }

        /// <summary>
        /// Clear all column data formats. This is effectively making all columns to be of General type.
        /// </summary>
        public void ClearColumnFormats()
        {
            this.dictColumnFormat.Clear();
        }

        /// <summary>
        /// Add custom date formats (this is the .NET date format code, not Excel's format code).
        /// This is used to parse any date data into a date, and is done first before trying other
        /// generic date parsing operations.
        /// </summary>
        /// <param name="Format"></param>
        public void AddCustomDateFormat(string Format)
        {
            this.listCustomDateFormats.Add(Format);
        }

        /// <summary>
        /// Clear all custom date formats.
        /// </summary>
        public void ClearCustomDateFormats()
        {
            this.listCustomDateFormats.Clear();
        }

        /// <summary>
        /// Set the width of a column in number of characters for separating data columns.
        /// This is used when the data source is specified as of fixed width.
        /// If no width is specified, the DefaultFixedWidth is used.
        /// </summary>
        /// <param name="ColumnIndex">The column index of the data source. This is 1-based indexing, so it's 1 for the 1st data source column, 2 for the 2nd data source column and so on.</param>
        /// <param name="ColumnWidth">The column width in number of characters.</param>
        public void SetFixedWidth(int ColumnIndex, int ColumnWidth)
        {
            if (ColumnIndex >= 1 && ColumnWidth >= 1)
            {
                this.dictFixedWidth[ColumnIndex] = ColumnWidth;
            }
        }

        /// <summary>
        /// Clone an instance of SLTextImportOptions.
        /// </summary>
        /// <returns>An SLTextImportOptions object.</returns>
        public SLTextImportOptions Clone()
        {
            SLTextImportOptions tio = new SLTextImportOptions();
            tio.DataFieldType = this.DataFieldType;
            tio.iDefaultFixedWidth = this.iDefaultFixedWidth;
            tio.UseTabDelimiter = this.UseTabDelimiter;
            tio.UseSemicolonDelimiter = this.UseSemicolonDelimiter;
            tio.UseCommaDelimiter = this.UseCommaDelimiter;
            tio.UseSpaceDelimiter = this.UseSpaceDelimiter;
            tio.UseCustomDelimiter = this.UseCustomDelimiter;
            tio.CustomDelimiter = this.CustomDelimiter;
            tio.MergeDelimiters = this.MergeDelimiters;
            tio.HasTextQualifier = this.HasTextQualifier;
            tio.TextQualifier = this.TextQualifier;
            tio.iImportStartRowIndex = this.iImportStartRowIndex;
            tio.Culture = this.Culture;
            tio.NumberStyles = this.NumberStyles;
            tio.Encoding = this.Encoding;
            tio.PreserveSpace = this.PreserveSpace;
            
            foreach (int key in this.dictColumnFormat.Keys)
            {
                tio.dictColumnFormat[key] = this.dictColumnFormat[key];
            }

            foreach (string format in this.listCustomDateFormats)
            {
                tio.listCustomDateFormats.Add(format);
            }

            foreach (int key in this.dictFixedWidth.Keys)
            {
                tio.dictFixedWidth[key] = this.dictFixedWidth[key];
            }

            return tio;
        }
    }
}
