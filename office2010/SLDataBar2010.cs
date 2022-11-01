using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    internal class SLDataBar2010
    {
        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.databar.aspx

        internal SLConditionalFormattingValueObject2010 Cfvo1 { get; set; }
        internal SLConditionalFormattingValueObject2010 Cfvo2 { get; set; }
        internal SLColor FillColor { get; set; }
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

        internal SLDataBar2010()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Cfvo1 = new SLConditionalFormattingValueObject2010();
            this.Cfvo2 = new SLConditionalFormattingValueObject2010();
            this.FillColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.BorderColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.NegativeFillColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.NegativeBorderColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.AxisColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());

            this.MinLength = 10;
            this.MaxLength = 90;
            this.ShowValue = true;
            this.Border = false;
            this.Gradient = true;
            this.Direction = X14.DataBarDirectionValues.Context;
            this.NegativeBarColorSameAsPositive = false;
            this.NegativeBarBorderColorSameAsPositive = true;
            this.AxisPosition = X14.DataBarAxisPositionValues.Automatic;
        }

        internal void FromDataBar(X14.DataBar db)
        {
            this.SetAllNull();

            using (OpenXmlReader oxr = OpenXmlReader.Create(db))
            {
                int i = 0;
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(X14.ConditionalFormattingValueObject))
                    {
                        if (i == 0)
                        {
                            this.Cfvo1.FromConditionalFormattingValueObject((X14.ConditionalFormattingValueObject)oxr.LoadCurrentElement());
                            ++i;
                        }
                        else if (i == 1)
                        {
                            this.Cfvo2.FromConditionalFormattingValueObject((X14.ConditionalFormattingValueObject)oxr.LoadCurrentElement());
                            ++i;
                        }
                    }
                    else if (oxr.ElementType == typeof(X14.FillColor))
                    {
                        this.FillColor.FromFillColor((X14.FillColor)oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(X14.BorderColor))
                    {
                        this.BorderColor.FromBorderColor((X14.BorderColor)oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(X14.NegativeFillColor))
                    {
                        this.NegativeFillColor.FromNegativeFillColor((X14.NegativeFillColor)oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(X14.NegativeBorderColor))
                    {
                        this.NegativeBorderColor.FromNegativeBorderColor((X14.NegativeBorderColor)oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(X14.BarAxisColor))
                    {
                        this.AxisColor.FromBarAxisColor((X14.BarAxisColor)oxr.LoadCurrentElement());
                    }
                }
            }

            if (db.MinLength != null) this.MinLength = db.MinLength.Value;
            if (db.MaxLength != null) this.MaxLength = db.MaxLength.Value;
            if (db.ShowValue != null) this.ShowValue = db.ShowValue.Value;
            if (db.Border != null) this.Border = db.Border.Value;
            if (db.Gradient != null) this.Gradient = db.Gradient.Value;
            if (db.Direction != null) this.Direction = db.Direction.Value;
            if (db.NegativeBarColorSameAsPositive != null) this.NegativeBarColorSameAsPositive = db.NegativeBarColorSameAsPositive.Value;
            if (db.NegativeBarBorderColorSameAsPositive != null) this.NegativeBarBorderColorSameAsPositive = db.NegativeBarBorderColorSameAsPositive.Value;
            if (db.AxisPosition != null) this.AxisPosition = db.AxisPosition.Value;
        }

        internal X14.DataBar ToDataBar(bool RenderFillColor)
        {
            X14.DataBar db = new X14.DataBar();
            if (this.MinLength != 10) db.MinLength = this.MinLength;

            // according to Open XML specs, this cannot be more than 100 percent.
            if (this.MaxLength > 100) this.MaxLength = 100;
            if (this.MaxLength != 90) db.MaxLength = this.MaxLength;

            if (!this.ShowValue) db.ShowValue = this.ShowValue;
            if (this.Border) db.Border = this.Border;
            if (!this.Gradient) db.Gradient = this.Gradient;
            if (this.Direction != X14.DataBarDirectionValues.Context) db.Direction = this.Direction;
            if (this.NegativeBarColorSameAsPositive) db.NegativeBarColorSameAsPositive = this.NegativeBarColorSameAsPositive;
            if (!this.NegativeBarBorderColorSameAsPositive) db.NegativeBarBorderColorSameAsPositive = this.NegativeBarBorderColorSameAsPositive;
            if (this.AxisPosition != X14.DataBarAxisPositionValues.Automatic) db.AxisPosition = this.AxisPosition;

            db.Append(this.Cfvo1.ToConditionalFormattingValueObject());
            db.Append(this.Cfvo2.ToConditionalFormattingValueObject());

            // The condition is mainly if the priority of the parent rule exists. See Open XML specs.
            if (RenderFillColor) db.Append(this.FillColor.ToFillColor());

            if (this.Border) db.Append(this.BorderColor.ToBorderColor());
            if (!this.NegativeBarColorSameAsPositive) db.Append(this.NegativeFillColor.ToNegativeFillColor());
            if (!this.NegativeBarBorderColorSameAsPositive && this.Border) db.Append(this.NegativeBorderColor.ToNegativeBorderColor());
            if (this.AxisPosition != X14.DataBarAxisPositionValues.None) db.Append(this.AxisColor.ToBarAxisColor());

            return db;
        }

        internal SLDataBar2010 Clone()
        {
            SLDataBar2010 db = new SLDataBar2010();
            db.Cfvo1 = this.Cfvo1.Clone();
            db.Cfvo2 = this.Cfvo2.Clone();
            db.FillColor = this.FillColor.Clone();
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
