using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    internal class SLDataBar
    {
        internal bool Is2010;

        internal SLConditionalFormatAutoMinMaxValues MinimumType { get; set; }
        internal string MinimumValue { get; set; }
        internal SLConditionalFormatAutoMinMaxValues MaximumType { get; set; }
        internal string MaximumValue { get; set; }

        internal SLColor Color { get; set; }
        internal SLColor BorderColor { get; set; }
        internal SLColor NegativeFillColor { get; set; }
        internal SLColor NegativeBorderColor { get; set; }
        internal SLColor AxisColor { get; set; }
        internal uint MinLength { get; set; }
        internal uint MaxLength { get; set; }
        internal bool ShowValue { get; set; }
        internal bool Border { get; set; }
        internal bool Gradient { get; set; }
        internal X14.DataBarDirectionValues Direction { get; set; }
        internal bool NegativeBarColorSameAsPositive { get; set; }
        internal bool NegativeBarBorderColorSameAsPositive { get; set; }
        internal X14.DataBarAxisPositionValues AxisPosition { get; set; }

        internal SLDataBar()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Is2010 = false;
            this.MinimumType = SLConditionalFormatAutoMinMaxValues.Percentile;
            this.MinimumValue = string.Empty;
            this.MaximumType = SLConditionalFormatAutoMinMaxValues.Percentile;
            this.MaximumValue = string.Empty;
            this.Color = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.BorderColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.NegativeFillColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.NegativeBorderColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.AxisColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.MinLength = 10;
            this.MaxLength = 90;
            this.ShowValue = true;
        }

        internal void FromDataBar(DataBar db)
        {
            this.SetAllNull();

            using (OpenXmlReader oxr = OpenXmlReader.Create(db))
            {
                int i = 0;
                SLConditionalFormatValueObject cfvo;
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(ConditionalFormatValueObject))
                    {
                        if (i == 0)
                        {
                            cfvo = new SLConditionalFormatValueObject();
                            cfvo.FromConditionalFormatValueObject((ConditionalFormatValueObject)oxr.LoadCurrentElement());
                            this.MinimumType = this.TranslateToAutoMinMaxValues(cfvo.Type);
                            this.MinimumValue = cfvo.Val;
                            ++i;
                        }
                        else if (i == 1)
                        {
                            cfvo = new SLConditionalFormatValueObject();
                            cfvo.FromConditionalFormatValueObject((ConditionalFormatValueObject)oxr.LoadCurrentElement());
                            this.MaximumType = this.TranslateToAutoMinMaxValues(cfvo.Type);
                            this.MaximumValue = cfvo.Val;
                            ++i;
                        }
                    }
                    else if (oxr.ElementType == typeof(Color))
                    {
                        this.Color.FromSpreadsheetColor((Color)oxr.LoadCurrentElement());
                    }
                }
            }

            if (db.MinLength != null) this.MinLength = db.MinLength.Value;
            if (db.MaxLength != null) this.MaxLength = db.MaxLength.Value;
            if (db.ShowValue != null) this.ShowValue = db.ShowValue.Value;
        }

        internal SLConditionalFormatAutoMinMaxValues TranslateToAutoMinMaxValues(ConditionalFormatValueObjectValues Type)
        {
            SLConditionalFormatAutoMinMaxValues result = SLConditionalFormatAutoMinMaxValues.Percentile;
            switch (Type)
            {
                case ConditionalFormatValueObjectValues.Formula:
                    result = SLConditionalFormatAutoMinMaxValues.Formula;
                    break;
                case ConditionalFormatValueObjectValues.Max:
                    result = SLConditionalFormatAutoMinMaxValues.Value;
                    break;
                case ConditionalFormatValueObjectValues.Min:
                    result = SLConditionalFormatAutoMinMaxValues.Value;
                    break;
                case ConditionalFormatValueObjectValues.Number:
                    result = SLConditionalFormatAutoMinMaxValues.Number;
                    break;
                case ConditionalFormatValueObjectValues.Percent:
                    result = SLConditionalFormatAutoMinMaxValues.Percent;
                    break;
                case ConditionalFormatValueObjectValues.Percentile:
                    result = SLConditionalFormatAutoMinMaxValues.Percentile;
                    break;
            }

            return result;
        }

        internal DataBar ToDataBar()
        {
            DataBar db = new DataBar();
            if (this.MinLength != 10) db.MinLength = this.MinLength;
            if (this.MaxLength != 90) db.MaxLength = this.MaxLength;
            if (!this.ShowValue) db.ShowValue = this.ShowValue;

            SLConditionalFormatValueObject cfvo;

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = ConditionalFormatValueObjectValues.Min;
            switch (this.MinimumType)
            {
                case SLConditionalFormatAutoMinMaxValues.Automatic:
                    cfvo.Type = ConditionalFormatValueObjectValues.Min;
                    cfvo.Val = string.Empty;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Formula:
                    cfvo.Type = ConditionalFormatValueObjectValues.Formula;
                    cfvo.Val = this.MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Number:
                    cfvo.Type = ConditionalFormatValueObjectValues.Number;
                    cfvo.Val = this.MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percent:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percent;
                    cfvo.Val = this.MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percentile:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percentile;
                    cfvo.Val = this.MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Value:
                    cfvo.Type = ConditionalFormatValueObjectValues.Min;
                    cfvo.Val = string.Empty;
                    break;
            }
            db.Append(cfvo.ToConditionalFormatValueObject());

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = ConditionalFormatValueObjectValues.Max;
            switch (this.MaximumType)
            {
                case SLConditionalFormatAutoMinMaxValues.Automatic:
                    cfvo.Type = ConditionalFormatValueObjectValues.Max;
                    cfvo.Val = string.Empty;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Formula:
                    cfvo.Type = ConditionalFormatValueObjectValues.Formula;
                    cfvo.Val = this.MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Number:
                    cfvo.Type = ConditionalFormatValueObjectValues.Number;
                    cfvo.Val = this.MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percent:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percent;
                    cfvo.Val = this.MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percentile:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percentile;
                    cfvo.Val = this.MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Value:
                    cfvo.Type = ConditionalFormatValueObjectValues.Max;
                    cfvo.Val = string.Empty;
                    break;
            }
            db.Append(cfvo.ToConditionalFormatValueObject());

            db.Append(this.Color.ToSpreadsheetColor());

            return db;
        }

        internal SLDataBar2010 ToDataBar2010()
        {
            SLDataBar2010 db = new SLDataBar2010();
            switch (this.MinimumType)
            {
                case SLConditionalFormatAutoMinMaxValues.Automatic:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.AutoMin;
                    db.Cfvo1.Formula = string.Empty;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Formula:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Formula;
                    db.Cfvo1.Formula = this.MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Number:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric;
                    db.Cfvo1.Formula = this.MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percent:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Percent;
                    db.Cfvo1.Formula = this.MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percentile:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Percentile;
                    db.Cfvo1.Formula = this.MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Value:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Min;
                    db.Cfvo1.Formula = string.Empty;
                    break;
            }

            switch (this.MaximumType)
            {
                case SLConditionalFormatAutoMinMaxValues.Automatic:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.AutoMax;
                    db.Cfvo2.Formula = string.Empty;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Formula:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Formula;
                    db.Cfvo2.Formula = this.MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Number:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric;
                    db.Cfvo2.Formula = this.MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percent:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Percent;
                    db.Cfvo2.Formula = this.MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percentile:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Percentile;
                    db.Cfvo2.Formula = this.MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Value:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Max;
                    db.Cfvo2.Formula = string.Empty;
                    break;
            }

            db.FillColor = this.Color.Clone();
            db.BorderColor = this.BorderColor.Clone();
            db.NegativeFillColor = this.NegativeFillColor.Clone();
            db.NegativeBorderColor = this.NegativeBorderColor.Clone();
            db.AxisColor = this.AxisColor.Clone();
            db.MinLength = this.MinLength;
            db.MaxLength = this.MaxLength;
            db.ShowValue = this.ShowValue;
            db.Border = this.Border;
            db.Gradient = this.Gradient;
            db.Direction = this.Direction;
            db.NegativeBarColorSameAsPositive = this.NegativeBarColorSameAsPositive;
            db.NegativeBarBorderColorSameAsPositive = this.NegativeBarBorderColorSameAsPositive;
            db.AxisPosition = this.AxisPosition;

            return db;
        }

        internal SLDataBar Clone()
        {
            SLDataBar db = new SLDataBar();
            db.Is2010 = this.Is2010;
            db.MinimumType = this.MinimumType;
            db.MinimumValue = this.MinimumValue;
            db.MaximumType = this.MaximumType;
            db.MaximumValue = this.MaximumValue;
            db.Color = this.Color.Clone();
            db.BorderColor = this.BorderColor.Clone();
            db.NegativeFillColor = this.NegativeFillColor.Clone();
            db.NegativeBorderColor = this.NegativeBorderColor.Clone();
            db.AxisColor = this.AxisColor.Clone();
            db.MinLength = this.MinLength;
            db.MaxLength = this.MaxLength;
            db.ShowValue = this.ShowValue;
            db.Border = this.Border;
            db.Gradient = this.Gradient;
            db.Direction = this.Direction;
            db.NegativeBarColorSameAsPositive = this.NegativeBarColorSameAsPositive;
            db.NegativeBarBorderColorSameAsPositive = this.NegativeBarBorderColorSameAsPositive;
            db.AxisPosition = this.AxisPosition;

            return db;
        }
    }
}
