using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting value axes in charts.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.ValueAxis class.
    /// </summary>
    public class SLValueAxis : EGAxShared
    {
        // the actual value is stored at the category/date/value axis
        internal C.CrossBetweenValues CrossBetween { get; set; }
        
        /// <summary>
        /// The major unit on the axis. A null value means it's automatically set.
        /// </summary>
        public double? MajorUnit { get; set; }

        /// <summary>
        /// The minor unit on the axis. A null value means it's automatically set.
        /// </summary>
        public double? MinorUnit { get; set; }

        /// <summary>
        /// Logarithmic scale of the axis, ranging from 2 to 1000 (both inclusive). A null value means it's not used.
        /// </summary>
        public double? LogarithmicScale
        {
            get { return this.LogBase; }
            set { this.LogBase = value; }
        }

        // C.DisplayUnits
        internal C.BuiltInUnitValues? BuiltInUnitValues { get; set; }
        internal bool ShowDisplayUnitsLabel { get; set; }

        /// <summary>
        /// The maximum value on the axis. A null value means it's automatically set.
        /// </summary>
        public double? Maximum
        {
            get { return this.MaxAxisValue; }
            set { this.MaxAxisValue = value; }
        }

        /// <summary>
        /// The minimum value on the axis. A null value means it's automatically set.
        /// </summary>
        public double? Minimum
        {
            get { return this.MinAxisValue; }
            set { this.MinAxisValue = value; }
        }

        internal SLValueAxis(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
            : base(ThemeColors, IsStylish, ThrowExceptionsIfAny)
        {
            this.CrossBetween = C.CrossBetweenValues.Between;
            this.MajorUnit = null;
            this.MinorUnit = null;
            this.BuiltInUnitValues = null;
            this.ShowDisplayUnitsLabel = false;

            if (IsStylish)
            {
                this.ShapeProperties.Fill.SetNoFill();
                this.ShapeProperties.Outline.SetNoLine();
            }
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
        }

        /// <summary>
        /// Set the display units on the axis.
        /// </summary>
        /// <param name="BuiltInUnit">Built-in unit types.</param>
        /// <param name="ShowDisplayUnitsLabel">True to show the display units label on the chart. False otherwise.</param>
        public void SetDisplayUnits(C.BuiltInUnitValues BuiltInUnit, bool ShowDisplayUnitsLabel)
        {
            this.BuiltInUnitValues = BuiltInUnit;
            this.ShowDisplayUnitsLabel = ShowDisplayUnitsLabel;
        }

        /// <summary>
        /// Remove the display units on the axis.
        /// </summary>
        public void RemoveDisplayUnits()
        {
            this.BuiltInUnitValues = null;
            this.ShowDisplayUnitsLabel = false;
        }

        /// <summary>
        /// Set the corresponding category/date/value axis to cross this axis at an automatic value.
        /// </summary>
        public void SetAutomaticOtherAxisCrossing()
        {
            this.OtherAxisIsCrosses = true;
            this.OtherAxisCrosses = C.CrossesValues.AutoZero;
            this.OtherAxisCrossesAt = 0;
        }

        /// <summary>
        /// Set the corresponding category/date/value axis to cross this axis at a given value.
        /// </summary>
        /// <param name="CrossingAxisValue">Axis value to cross at.</param>
        public void SetOtherAxisCrossing(double CrossingAxisValue)
        {
            this.OtherAxisIsCrosses = false;
            this.OtherAxisCrosses = C.CrossesValues.AutoZero;
            this.OtherAxisCrossesAt = CrossingAxisValue;
        }

        /// <summary>
        /// Set the corresponding category/date/value axis to cross this axis at the maximum value.
        /// </summary>
        public void SetMaximumOtherAxisCrossing()
        {
            this.OtherAxisIsCrosses = true;
            this.OtherAxisCrosses = C.CrossesValues.Maximum;
            this.OtherAxisCrossesAt = 0;
        }

        internal C.ValueAxis ToValueAxis(bool IsStylish)
        {
            C.ValueAxis va = new C.ValueAxis();
            va.AxisId = new C.AxisId() { Val = this.AxisId };

            va.Scaling = new C.Scaling();
            va.Scaling.Orientation = new C.Orientation() { Val = this.Orientation };
            if (this.LogBase != null) va.Scaling.LogBase = new C.LogBase() { Val = this.LogBase.Value };
            if (this.MaxAxisValue != null) va.Scaling.MaxAxisValue = new C.MaxAxisValue() { Val = this.MaxAxisValue.Value };
            if (this.MinAxisValue != null) va.Scaling.MinAxisValue = new C.MinAxisValue() { Val = this.MinAxisValue.Value };

            va.Delete = new C.Delete() { Val = this.Delete };

            C.AxisPositionValues axpos = this.AxisPosition;
            if (!this.ForceAxisPosition)
            {
                if (this.OtherAxisIsInReverseOrder) axpos = SLChartTool.GetOppositePosition(axpos);
                if (this.OtherAxisCrossedAtMaximum) axpos = SLChartTool.GetOppositePosition(axpos);
            }
            va.AxisPosition = new C.AxisPosition() { Val = axpos };

            if (this.ShowMajorGridlines)
            {
                va.MajorGridlines = this.MajorGridlines.ToMajorGridlines(IsStylish);
            }

            if (this.ShowMinorGridlines)
            {
                va.MinorGridlines = this.MinorGridlines.ToMinorGridlines(IsStylish);
            }

            if (this.ShowTitle)
            {
                va.Title = this.Title.ToTitle(IsStylish);
            }

            if (this.HasNumberingFormat)
            {
                va.NumberingFormat = new C.NumberingFormat()
                {
                    FormatCode = this.FormatCode,
                    SourceLinked = this.SourceLinked
                };
            }

            va.MajorTickMark = new C.MajorTickMark() { Val = this.MajorTickMark };
            va.MinorTickMark = new C.MinorTickMark() { Val = this.MinorTickMark };
            va.TickLabelPosition = new C.TickLabelPosition() { Val = this.TickLabelPosition };

            if (this.ShapeProperties.HasShapeProperties) va.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            if (this.Rotation != null || this.Vertical != null || this.Anchor != null || this.AnchorCenter != null)
            {
                va.TextProperties = new C.TextProperties();
                va.TextProperties.BodyProperties = new A.BodyProperties();
                if (this.Rotation != null) va.TextProperties.BodyProperties.Rotation = (int)(this.Rotation.Value * SLConstants.DegreeToAngleRepresentation);
                if (this.Vertical != null) va.TextProperties.BodyProperties.Vertical = this.Vertical.Value;
                if (this.Anchor != null) va.TextProperties.BodyProperties.Anchor = this.Anchor.Value;
                if (this.AnchorCenter != null) va.TextProperties.BodyProperties.AnchorCenter = this.AnchorCenter.Value;

                va.TextProperties.ListStyle = new A.ListStyle();

                A.Paragraph para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();
                para.ParagraphProperties.Append(new A.DefaultRunProperties());
                va.TextProperties.Append(para);
            }
            else if (IsStylish)
            {
                va.TextProperties = new C.TextProperties();
                va.TextProperties.BodyProperties = new A.BodyProperties()
                {
                    Rotation = -60000000,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                va.TextProperties.ListStyle = new A.ListStyle();

                A.Paragraph para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();

                A.DefaultRunProperties defrunprops = new A.DefaultRunProperties();
                defrunprops.FontSize = 900;
                defrunprops.Bold = false;
                defrunprops.Italic = false;
                defrunprops.Underline = A.TextUnderlineValues.None;
                defrunprops.Strike = A.TextStrikeValues.NoStrike;
                defrunprops.Kerning = 1200;
                defrunprops.Baseline = 0;

                A.SchemeColor schclr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
                schclr.Append(new A.LuminanceModulation() { Val = 65000 });
                schclr.Append(new A.LuminanceOffset() { Val = 35000 });
                defrunprops.Append(new A.SolidFill()
                {
                    SchemeColor = schclr
                });

                defrunprops.Append(new A.LatinFont() { Typeface = "+mn-lt" });
                defrunprops.Append(new A.EastAsianFont() { Typeface = "+mn-ea" });
                defrunprops.Append(new A.ComplexScriptFont() { Typeface = "+mn-cs" });

                para.ParagraphProperties.Append(defrunprops);
                para.Append(new A.EndParagraphRunProperties() { Language = System.Globalization.CultureInfo.CurrentCulture.Name });

                va.TextProperties.Append(para);
            }

            va.CrossingAxis = new C.CrossingAxis() { Val = this.CrossingAxis };

            if (this.IsCrosses != null)
            {
                if (this.IsCrosses.Value)
                {
                    va.Append(new C.Crosses() { Val = this.Crosses });
                }
                else
                {
                    va.Append(new C.CrossesAt() { Val = this.CrossesAt });
                }
            }

            va.Append(new C.CrossBetween() { Val = this.CrossBetween });
            if (this.MajorUnit != null) va.Append(new C.MajorUnit() { Val = this.MajorUnit.Value });
            if (this.MinorUnit != null) va.Append(new C.MinorUnit() { Val = this.MinorUnit.Value });

            if (this.BuiltInUnitValues != null)
            {
                C.DisplayUnits du = new C.DisplayUnits();
                du.Append(new C.BuiltInUnit() { Val = this.BuiltInUnitValues.Value });
                if (this.ShowDisplayUnitsLabel)
                {
                    C.DisplayUnitsLabel dul = new C.DisplayUnitsLabel();
                    dul.Layout = new C.Layout();
                    du.Append(dul);
                }
                va.Append(du);
            }

            return va;
        }

        internal SLValueAxis Clone()
        {
            SLValueAxis va = new SLValueAxis(this.ShapeProperties.listThemeColors, false, this.ShapeProperties.ThrowExceptionsIfAny);
            va.Rotation = this.Rotation;
            va.Vertical = this.Vertical;
            va.Anchor = this.Anchor;
            va.AnchorCenter = this.AnchorCenter;
            va.AxisId = this.AxisId;
            va.fLogBase = this.fLogBase;
            va.Orientation = this.Orientation;
            va.MaxAxisValue = this.MaxAxisValue;
            va.MinAxisValue = this.MinAxisValue;
            va.OtherAxisIsInReverseOrder = this.OtherAxisIsInReverseOrder;
            va.OtherAxisCrossedAtMaximum = this.OtherAxisCrossedAtMaximum;
            va.Delete = this.Delete;
            va.ForceAxisPosition = this.ForceAxisPosition;
            va.AxisPosition = this.AxisPosition;
            va.ShowMajorGridlines = this.ShowMajorGridlines;
            va.MajorGridlines = this.MajorGridlines.Clone();
            va.ShowMinorGridlines = this.ShowMinorGridlines;
            va.MinorGridlines = this.MinorGridlines.Clone();
            va.ShowTitle = this.ShowTitle;
            va.Title = this.Title.Clone();
            va.HasNumberingFormat = this.HasNumberingFormat;
            va.sFormatCode = this.sFormatCode;
            va.bSourceLinked = this.bSourceLinked;
            va.MajorTickMark = this.MajorTickMark;
            va.MinorTickMark = this.MinorTickMark;
            va.TickLabelPosition = this.TickLabelPosition;
            va.ShapeProperties = this.ShapeProperties.Clone();
            va.CrossingAxis = this.CrossingAxis;
            va.IsCrosses = this.IsCrosses;
            va.Crosses = this.Crosses;
            va.CrossesAt = this.CrossesAt;
            va.OtherAxisIsCrosses = this.OtherAxisIsCrosses;
            va.OtherAxisCrosses = this.OtherAxisCrosses;
            va.OtherAxisCrossesAt = this.OtherAxisCrossesAt;

            va.CrossBetween = this.CrossBetween;
            va.MajorUnit = this.MajorUnit;
            va.MinorUnit = this.MinorUnit;
            va.BuiltInUnitValues = this.BuiltInUnitValues;
            va.ShowDisplayUnitsLabel = this.ShowDisplayUnitsLabel;

            return va;
        }
    }
}
