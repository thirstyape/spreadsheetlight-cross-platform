using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Data validation types.
    /// </summary>
    public enum SLDataValidationAllowedValues
    {
        /// <summary>
        /// Whole number.
        /// </summary>
        WholeNumber = 0,
        /// <summary>
        /// Decimal.
        /// </summary>
        Decimal,
        /// <summary>
        /// Date.
        /// </summary>
        Date,
        /// <summary>
        /// Time.
        /// </summary>
        Time,
        /// <summary>
        /// Text length.
        /// </summary>
        TextLength
    }

    /// <summary>
    /// Data validation operations with 1 operand.
    /// </summary>
    public enum SLDataValidationSingleOperandValues
    {
        /// <summary>
        /// Equal.
        /// </summary>
        Equal = 0,
        /// <summary>
        /// Not equal.
        /// </summary>
        NotEqual,
        /// <summary>
        /// Greater than.
        /// </summary>
        GreaterThan,
        /// <summary>
        /// Less than.
        /// </summary>
        LessThan,
        /// <summary>
        /// Greater than or equal.
        /// </summary>
        GreaterThanOrEqual,
        /// <summary>
        /// Less than or equal.
        /// </summary>
        LessThanOrEqual
    }

    /// <summary>
    /// Encapsulates properties and methods for data validations.
    /// </summary>
    public class SLDataValidation
    {
        internal bool Date1904 { get; set; }

        internal string Formula1 { get; set; }
        internal string Formula2 { get; set; }

        internal bool HasDataValidation
        {
            get
            {
                return this.SequenceOfReferences.Count > 0 && (this.Type != DataValidationValues.None
                    || this.ErrorTitle.Length > 0 || this.Error.Length > 0
                    || this.PromptTitle.Length > 0 || this.Prompt.Length > 0);
            }
        }

        internal DataValidationValues Type { get; set; }
        internal DataValidationErrorStyleValues ErrorStyle { get; set; }
        internal DataValidationImeModeValues ImeMode { get; set; }
        internal DataValidationOperatorValues Operator { get; set; }

        internal bool AllowBlank { get; set; }
        internal bool ShowDropDown { get; set; }

        /// <summary>
        /// Specifies if the input message is shown.
        /// </summary>
        public bool ShowInputMessage { get; set; }

        /// <summary>
        /// Specifies if the error message is shown.
        /// </summary>
        public bool ShowErrorMessage { get; set; }

        internal string ErrorTitle { get; set; }
        internal string Error { get; set; }
        internal string PromptTitle { get; set; }
        internal string Prompt { get; set; }

        internal List<SLCellPointRange> SequenceOfReferences { get; set; }

        internal SLDataValidation()
        {
        }

        internal void InitialiseDataValidation(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool Date1904)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            this.SetAllNull();
            this.Date1904 = Date1904;
            this.SequenceOfReferences.Add(new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex));
        }

        private void SetAllNull()
        {
            this.Date1904 = false;
            this.Formula1 = string.Empty;
            this.Formula2 = string.Empty;
            this.Type = DataValidationValues.None;
            this.ErrorStyle = DataValidationErrorStyleValues.Stop;
            this.ImeMode = DataValidationImeModeValues.NoControl;
            this.Operator = DataValidationOperatorValues.Between;
            this.AllowBlank = false;
            this.ShowDropDown = false;
            this.ShowInputMessage = true;
            this.ShowErrorMessage = true;
            this.ErrorTitle = string.Empty;
            this.Error = string.Empty;
            this.PromptTitle = string.Empty;
            this.Prompt = string.Empty;
            this.SequenceOfReferences = new List<SLCellPointRange>();
        }

        /// <summary>
        /// Allow any value.
        /// </summary>
        public void AllowAnyValue()
        {
            this.Type = DataValidationValues.None;
        }

        /// <summary>
        /// Allow only whole numbers.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(bool IsBetween, int Minimum, int Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Whole;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow only whole numbers.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(bool IsBetween, long Minimum, long Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Whole;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow only whole numbers.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(bool IsBetween, string Minimum, string Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Whole;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = this.CleanDataSourceForFormula(Minimum);
            this.Formula2 = this.CleanDataSourceForFormula(Maximum);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow only whole numbers.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(SLDataValidationSingleOperandValues DataOperator, int DataValue, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Whole;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow only whole numbers.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(SLDataValidationSingleOperandValues DataOperator, long DataValue, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Whole;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow only whole numbers.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowWholeNumber(SLDataValidationSingleOperandValues DataOperator, string DataValue, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Whole;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = this.CleanDataSourceForFormula(DataValue);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow decimal (floating point) values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(bool IsBetween, float Minimum, float Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Decimal;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow decimal (floating point) values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(bool IsBetween, double Minimum, double Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Decimal;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow decimal (floating point) values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(bool IsBetween, decimal Minimum, decimal Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Decimal;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow decimal (floating point) values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(bool IsBetween, string Minimum, string Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Decimal;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = this.CleanDataSourceForFormula(Minimum);
            this.Formula2 = this.CleanDataSourceForFormula(Maximum);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow decimal (floating point) values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(SLDataValidationSingleOperandValues DataOperator, float DataValue, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Decimal;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow decimal (floating point) values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(SLDataValidationSingleOperandValues DataOperator, double DataValue, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Decimal;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow decimal (floating point) values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(SLDataValidationSingleOperandValues DataOperator, decimal DataValue, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Decimal;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = DataValue.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow decimal (floating point) values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDecimal(SLDataValidationSingleOperandValues DataOperator, string DataValue, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Decimal;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = this.CleanDataSourceForFormula(DataValue);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow a list of values.
        /// </summary>
        /// <param name="DataSource">The data source. For example, "$A$1:$A$5"</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        /// <param name="InCellDropDown">True if a dropdown list appears for selecting. False otherwise.</param>
        public void AllowList(string DataSource, bool IgnoreBlank, bool InCellDropDown)
        {
            this.Type = DataValidationValues.List;
            this.Operator = DataValidationOperatorValues.Between;

            if (DataSource.StartsWith("="))
            {
                this.Formula1 = DataSource.Substring(1);
            }
            else
            {
                this.Formula1 = DataSource;

                // NOTE (9 Dec 2016): what was I thinking? Normally the data source won't be weird...
                // Thanks to Ryan B for emailing in!
                //if (Regex.IsMatch(DataSource, "^\\s*\\$[A-Za-z]{1,3}\\$[0-9]{1,7}"))
                //{
                //    this.Formula1 = DataSource;
                //}
                //else
                //{
                //    // data source is something like 1,2,3
                //    // we need to make it "1,2,3"
                //    this.Formula1 = string.Format("\"{0}\"", DataSource);
                //}
            }

            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
            // I don't know why it's reversed. It seems to make sense when "normal"...
            this.ShowDropDown = !InCellDropDown;
        }

        /// <summary>
        /// Allow date values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDate(bool IsBetween, DateTime Minimum, DateTime Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Date;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = SLTool.CalculateDaysFromEpoch(Minimum, this.Date1904).ToString(CultureInfo.InvariantCulture);
            this.Formula2 = SLTool.CalculateDaysFromEpoch(Maximum, this.Date1904).ToString(CultureInfo.InvariantCulture);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow date values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value. Any valid date formatted value is fine. It is suggested to just copy the value you have in Excel interface.</param>
        /// <param name="Maximum">The maximum value. Any valid date formatted value is fine. It is suggested to just copy the value you have in Excel interface.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDate(bool IsBetween, string Minimum, string Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Date;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;

            DateTime dt;

            if (Minimum.StartsWith("="))
            {
                this.Formula1 = Minimum.Substring(1);
            }
            else
            {
                if (DateTime.TryParse(Minimum, out dt))
                {
                    this.Formula1 = SLTool.CalculateDaysFromEpoch(dt, this.Date1904).ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    // 1 Jan 1900
                    this.Formula1 = "1";
                }
            }

            if (Maximum.StartsWith("="))
            {
                this.Formula2 = Maximum.Substring(1);
            }
            else
            {
                if (DateTime.TryParse(Maximum, out dt))
                {
                    this.Formula2 = SLTool.CalculateDaysFromEpoch(dt, this.Date1904).ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    // 1 Jan 1900
                    this.Formula2 = "1";
                }
            }

            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow date values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value. Any valid date formatted value is fine. It is suggested to just copy the value you have in Excel interface.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDate(SLDataValidationSingleOperandValues DataOperator, DateTime DataValue, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Date;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = SLTool.CalculateDaysFromEpoch(DataValue, this.Date1904).ToString(CultureInfo.InvariantCulture);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow date values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="DataValue">The data value. Any valid date formatted value is fine. It is suggested to just copy the value you have in Excel interface.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowDate(SLDataValidationSingleOperandValues DataOperator, string DataValue, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Date;
            this.Operator = this.TranslateOperatorValues(DataOperator);

            DateTime dt;

            if (DataValue.StartsWith("="))
            {
                this.Formula1 = DataValue.Substring(1);
            }
            else
            {
                if (DateTime.TryParse(DataValue, out dt))
                {
                    this.Formula1 = SLTool.CalculateDaysFromEpoch(dt, this.Date1904).ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    // 1 Jan 1900
                    this.Formula1 = "1";
                }
            }

            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow time values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="StartHour">The start hour between 0 to 23 (both inclusive).</param>
        /// <param name="StartMinute">The start minute between 0 to 59 (both inclusive).</param>
        /// <param name="StartSecond">The start second between 0 to 59 (both inclusive).</param>
        /// <param name="EndHour">The end hour between 0 to 23 (both inclusive).</param>
        /// <param name="EndMinute">The end minute between 0 to 59 (both inclusive).</param>
        /// <param name="EndSecond">The end second between 0 to 59 (both inclusive).</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTime(bool IsBetween, int StartHour, int StartMinute, int StartSecond, int EndHour, int EndMinute, int EndSecond, bool IgnoreBlank)
        {
            if (StartHour < 0) StartHour = 0;
            if (StartHour > 23) StartHour = 23;
            if (StartMinute < 0) StartMinute = 0;
            if (StartMinute > 59) StartMinute = 59;
            if (StartSecond < 0) StartSecond = 0;
            if (StartSecond > 59) StartSecond = 59;
            if (EndHour < 0) EndHour = 0;
            if (EndHour > 23) EndHour = 23;
            if (EndMinute < 0) EndMinute = 0;
            if (EndMinute > 59) EndMinute = 59;
            if (EndSecond < 0) EndSecond = 0;
            if (EndSecond > 59) EndSecond = 59;

            this.Type = DataValidationValues.Time;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;

            double fTime = 0;

            // 1440 = 24 hours * 60 minutes
            // 86400 = 24 hours * 60 minutes * 60 seconds

            fTime = ((double)StartHour / 24.0) + ((double)StartMinute / 1440.0) + ((double)StartSecond / 86400.0);
            this.Formula1 = fTime.ToString(CultureInfo.InvariantCulture);

            fTime = ((double)EndHour / 24.0) + ((double)EndMinute / 1440.0) + ((double)EndSecond / 86400.0);
            this.Formula2 = fTime.ToString(CultureInfo.InvariantCulture);
            
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow time values.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="StartTime">The start time. Any valid time formatted value is fine. It is suggested to just copy the value you have in Excel interface.</param>
        /// <param name="EndTime">The end time. Any valid time formatted value is fine. It is suggested to just copy the value you have in Excel interface.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTime(bool IsBetween, string StartTime, string EndTime, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Time;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;

            double fTime = 0;
            DateTime dt;
            string sTime;
            // we include the day, month and year for formatting because it seems that parsing based
            // just on time (hour, minute, second, AM/PM designator) is too much for TryParseExact()...
            string[] saFormats = new string[] { "dd/MM/yyyy H", "dd/MM/yyyy h t", "dd/MM/yyyy h tt", "dd/MM/yyyy H:m", "dd/MM/yyyy h:m t", "dd/MM/yyyy h:m tt", "dd/MM/yyyy H:m:s", "dd/MM/yyyy h:m:s t", "dd/MM/yyyy h:m:s tt" };
            string sSampleDate = DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

            // 1440 = 24 hours * 60 minutes
            // 86400 = 24 hours * 60 minutes * 60 seconds

            if (StartTime.StartsWith("="))
            {
                this.Formula1 = StartTime.Substring(1);
            }
            else
            {
                sTime = string.Format("{0} {1}", sSampleDate, StartTime);
                if (DateTime.TryParseExact(sTime, saFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out dt))
                {
                    fTime = ((double)dt.Hour / 24.0) + ((double)dt.Minute / 1440.0) + ((double)dt.Second / 86400.0);
                    this.Formula1 = fTime.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    this.Formula1 = "0";
                }
            }

            if (EndTime.StartsWith("="))
            {
                this.Formula2 = EndTime.Substring(1);
            }
            else
            {
                sTime = string.Format("{0} {1}", sSampleDate, EndTime);
                if (DateTime.TryParseExact(sTime, saFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out dt))
                {
                    fTime = ((double)dt.Hour / 24.0) + ((double)dt.Minute / 1440.0) + ((double)dt.Second / 86400.0);
                    this.Formula1 = fTime.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    this.Formula1 = "0";
                }
            }

            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow time values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="Hour">The hour between 0 to 23 (both inclusive).</param>
        /// <param name="Minute">The minute between 0 to 59 (both inclusive).</param>
        /// <param name="Second">The second between 0 to 59 (both inclusive).</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTime(SLDataValidationSingleOperandValues DataOperator, int Hour, int Minute, int Second, bool IgnoreBlank)
        {
            if (Hour < 0) Hour = 0;
            if (Hour > 23) Hour = 23;
            if (Minute < 0) Minute = 0;
            if (Minute > 59) Minute = 59;
            if (Second < 0) Second = 0;
            if (Second > 59) Second = 59;

            this.Type = DataValidationValues.Time;
            this.Operator = this.TranslateOperatorValues(DataOperator);

            double fTime = 0;

            // 1440 = 24 hours * 60 minutes
            // 86400 = 24 hours * 60 minutes * 60 seconds

            fTime = ((double)Hour / 24.0) + ((double)Minute / 1440.0) + ((double)Second / 86400.0);
            this.Formula1 = fTime.ToString(CultureInfo.InvariantCulture);

            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow time values.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="Time">The time. Any valid time formatted value is fine. It is suggested to just copy the value you have in Excel interface.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTime(SLDataValidationSingleOperandValues DataOperator, string Time, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Time;
            this.Operator = this.TranslateOperatorValues(DataOperator);

            double fTime = 0;
            DateTime dt;
            string sTime;
            // we include the day, month and year for formatting because it seems that parsing based
            // just on time (hour, minute, second, AM/PM designator) is too much for TryParseExact()...
            string[] saFormats = new string[] { "dd/MM/yyyy H", "dd/MM/yyyy h t", "dd/MM/yyyy h tt", "dd/MM/yyyy H:m", "dd/MM/yyyy h:m t", "dd/MM/yyyy h:m tt", "dd/MM/yyyy H:m:s", "dd/MM/yyyy h:m:s t", "dd/MM/yyyy h:m:s tt" };
            string sSampleDate = DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);

            // 1440 = 24 hours * 60 minutes
            // 86400 = 24 hours * 60 minutes * 60 seconds

            if (Time.StartsWith("="))
            {
                this.Formula1 = Time.Substring(1);
            }
            else
            {
                sTime = string.Format("{0} {1}", sSampleDate, Time);
                if (DateTime.TryParseExact(sTime, saFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out dt))
                {
                    fTime = ((double)dt.Hour / 24.0) + ((double)dt.Minute / 1440.0) + ((double)dt.Second / 86400.0);
                    this.Formula1 = fTime.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    this.Formula1 = "0";
                }
            }

            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow data according to text length.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTextLength(bool IsBetween, int Minimum, int Maximum, bool IgnoreBlank)
        {
            if (Minimum < 0) Minimum = 0;
            if (Maximum < 0) Maximum = 0;

            this.Type = DataValidationValues.TextLength;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = Minimum.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = Maximum.ToString(CultureInfo.InvariantCulture);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow data according to text length.
        /// </summary>
        /// <param name="IsBetween">True if the data is between 2 values. False otherwise.</param>
        /// <param name="Minimum">The minimum value.</param>
        /// <param name="Maximum">The maximum value.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTextLength(bool IsBetween, string Minimum, string Maximum, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.TextLength;
            this.Operator = IsBetween ? DataValidationOperatorValues.Between : DataValidationOperatorValues.NotBetween;
            this.Formula1 = this.CleanDataSourceForFormula(Minimum);
            this.Formula2 = this.CleanDataSourceForFormula(Maximum);
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow data according to text length.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="Length">The text length for comparison.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTextLength(SLDataValidationSingleOperandValues DataOperator, int Length, bool IgnoreBlank)
        {
            if (Length < 0) Length = 0;

            this.Type = DataValidationValues.TextLength;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = Length.ToString(CultureInfo.InvariantCulture);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow data according to text length.
        /// </summary>
        /// <param name="DataOperator">The type of operation.</param>
        /// <param name="Length">The text length for comparison.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowTextLength(SLDataValidationSingleOperandValues DataOperator, string Length, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.TextLength;
            this.Operator = this.TranslateOperatorValues(DataOperator);
            this.Formula1 = this.CleanDataSourceForFormula(Length);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Allow custom validation.
        /// </summary>
        /// <param name="Formula">The formula used for validation.</param>
        /// <param name="IgnoreBlank">True if blanks are ignored. False otherwise.</param>
        public void AllowCustom(string Formula, bool IgnoreBlank)
        {
            this.Type = DataValidationValues.Custom;
            this.Operator = DataValidationOperatorValues.Between;
            this.Formula1 = this.CleanDataSourceForFormula(Formula);
            this.Formula2 = string.Empty;
            this.AllowBlank = IgnoreBlank;
        }

        /// <summary>
        /// Set the input message.
        /// </summary>
        /// <param name="Title">The title of the input message.</param>
        /// <param name="Message">The input message.</param>
        public void SetInputMessage(string Title, string Message)
        {
            this.PromptTitle = Title;
            this.Prompt = Message;
        }

        /// <summary>
        /// Set the error alert.
        /// </summary>
        /// <param name="ErrorStyle">The error style.</param>
        /// <param name="Title">The title of the error alert.</param>
        /// <param name="Message">The error message.</param>
        public void SetErrorAlert(DataValidationErrorStyleValues ErrorStyle, string Title, string Message)
        {
            this.ErrorStyle = ErrorStyle;
            this.ErrorTitle = Title;
            this.Error = Message;
        }

        /// <summary>
        /// Set the error alert.
        /// </summary>
        /// <param name="Title">The title of the error alert.</param>
        /// <param name="Message">The error message.</param>
        public void SetErrorAlert(string Title, string Message)
        {
            this.ErrorStyle = DataValidationErrorStyleValues.Stop;
            this.ErrorTitle = Title;
            this.Error = Message;
        }

        internal string CleanDataSourceForFormula(string DataValue)
        {
            string result = DataValue;
            if (result.StartsWith("="))
            {
                result = DataValue.Substring(1);
            }

            return result;
        }

        internal DataValidationOperatorValues TranslateOperatorValues(SLDataValidationSingleOperandValues Operator)
        {
            DataValidationOperatorValues result = DataValidationOperatorValues.Between;
            switch (Operator)
            {
                case SLDataValidationSingleOperandValues.Equal:
                    result = DataValidationOperatorValues.Equal;
                    break;
                case SLDataValidationSingleOperandValues.NotEqual:
                    result = DataValidationOperatorValues.NotEqual;
                    break;
                case SLDataValidationSingleOperandValues.GreaterThan:
                    result = DataValidationOperatorValues.GreaterThan;
                    break;
                case SLDataValidationSingleOperandValues.LessThan:
                    result = DataValidationOperatorValues.LessThan;
                    break;
                case SLDataValidationSingleOperandValues.GreaterThanOrEqual:
                    result = DataValidationOperatorValues.GreaterThanOrEqual;
                    break;
                case SLDataValidationSingleOperandValues.LessThanOrEqual:
                    result = DataValidationOperatorValues.LessThanOrEqual;
                    break;
            }

            return result;
        }

        internal void FromDataValidation(DataValidation dv)
        {
            this.SetAllNull();

            if (dv.Formula1 != null) this.Formula1 = dv.Formula1.Text;
            if (dv.Formula2 != null) this.Formula2 = dv.Formula2.Text;

            if (dv.Type != null) this.Type = dv.Type.Value;
            if (dv.ErrorStyle != null) this.ErrorStyle = dv.ErrorStyle.Value;
            if (dv.ImeMode != null) this.ImeMode = dv.ImeMode.Value;
            if (dv.Operator != null) this.Operator = dv.Operator.Value;
            if (dv.AllowBlank != null) this.AllowBlank = dv.AllowBlank.Value;
            if (dv.ShowDropDown != null) this.ShowDropDown = dv.ShowDropDown.Value;
            if (dv.ShowInputMessage != null) this.ShowInputMessage = dv.ShowInputMessage.Value;
            if (dv.ShowErrorMessage != null) this.ShowErrorMessage = dv.ShowErrorMessage.Value;

            if (dv.ErrorTitle != null) this.ErrorTitle = dv.ErrorTitle.Value;
            if (dv.Error != null) this.Error = dv.Error.Value;
            if (dv.PromptTitle != null) this.PromptTitle = dv.PromptTitle.Value;
            if (dv.Prompt != null) this.Prompt = dv.Prompt.Value;

            // it has to be not-null because it's a required thing, but you never know...
            if (dv.SequenceOfReferences != null)
            {
                this.SequenceOfReferences = SLTool.TranslateSeqRefToCellPointRange(dv.SequenceOfReferences);
            }
        }

        internal DataValidation ToDataValidation()
        {
            DataValidation dv = new DataValidation();

            if (this.Formula1.Length > 0) dv.Formula1 = new Formula1(this.Formula1);
            if (this.Formula2.Length > 0) dv.Formula2 = new Formula2(this.Formula2);

            if (this.Type != DataValidationValues.None) dv.Type = this.Type;
            if (this.ErrorStyle != DataValidationErrorStyleValues.Stop) dv.ErrorStyle = this.ErrorStyle;
            if (this.ImeMode != DataValidationImeModeValues.NoControl) dv.ImeMode = this.ImeMode;
            if (this.Operator != DataValidationOperatorValues.Between) dv.Operator = this.Operator;

            if (this.AllowBlank) dv.AllowBlank = this.AllowBlank;
            if (this.ShowDropDown) dv.ShowDropDown = this.ShowDropDown;
            if (this.ShowInputMessage) dv.ShowInputMessage = this.ShowInputMessage;
            if (this.ShowErrorMessage) dv.ShowErrorMessage = this.ShowErrorMessage;

            if (this.ErrorTitle.Length > 0) dv.ErrorTitle = this.ErrorTitle;
            if (this.Error.Length > 0) dv.Error = this.Error;

            if (this.PromptTitle.Length > 0) dv.PromptTitle = this.PromptTitle;
            if (this.Prompt.Length > 0) dv.Prompt = this.Prompt;

            dv.SequenceOfReferences = SLTool.TranslateCellPointRangeToSeqRef(this.SequenceOfReferences);

            return dv;
        }

        internal SLDataValidation Clone()
        {
            SLDataValidation dv = new SLDataValidation();
            dv.Date1904 = this.Date1904;
            dv.Formula1 = this.Formula1;
            dv.Formula2 = this.Formula2;
            dv.Type = this.Type;
            dv.ErrorStyle = this.ErrorStyle;
            dv.ImeMode = this.ImeMode;
            dv.Operator = this.Operator;
            dv.AllowBlank = this.AllowBlank;
            dv.ShowDropDown = this.ShowDropDown;
            dv.ShowInputMessage = this.ShowInputMessage;
            dv.ShowErrorMessage = this.ShowErrorMessage;
            dv.ErrorTitle = this.ErrorTitle;
            dv.Error = this.Error;
            dv.PromptTitle = this.PromptTitle;
            dv.Prompt = this.Prompt;

            dv.SequenceOfReferences = new List<SLCellPointRange>();
            foreach (SLCellPointRange pt in this.SequenceOfReferences)
            {
                dv.SequenceOfReferences.Add(pt);
            }

            return dv;
        }
    }
}
