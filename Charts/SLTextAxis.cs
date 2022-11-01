using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    internal enum SLAxisType
    {
        Category,
        Date,
        Value
    }

    /// <summary>
    /// Encapsulates properties and methods for setting chart axes, specifically simulating
    /// DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis,
    /// DocumentFormat.OpenXml.Drawing.Charts.DateAxis and
    /// DocumentFormat.OpenXml.Drawing.Charts.ValueAxis classes.
    /// </summary>
    public class SLTextAxis : EGAxShared
    {
        internal bool Date1904 { get; set; }

        internal SLAxisType AxisType { get; set; }

        // switch when axis types are changed (category to date or date to category)
        internal bool AutoLabeled { get; set; }

        internal ushort iTickLabelSkip;
        /// <summary>
        /// This is the interval between labels, and is at least 1. A suggested range is 1 to 255 (both inclusive). This is only for category axes.
        /// </summary>
        public ushort TickLabelSkip
        {
            get { return iTickLabelSkip; }
            set
            {
                iTickLabelSkip = value;
                if (iTickLabelSkip < 1) iTickLabelSkip = 1;
            }
        }

        internal ushort iTickMarkSkip;
        /// <summary>
        /// This is the interval between tick marks, and is at least 1. A suggested range is 1 to 31999 (both inclusive). This is only for category axes.
        /// </summary>
        public ushort TickMarkSkip
        {
            get { return iTickMarkSkip; }
            set
            {
                iTickMarkSkip = value;
                if (iTickMarkSkip < 1) iTickMarkSkip = 1;
            }
        }

        /// <summary>
        /// Label alignment for the category axis. This is ignored for date axes.
        /// </summary>
        public C.LabelAlignmentValues LabelAlignment { get; set; }

        internal ushort iLabelOffset;
        /// <summary>
        /// This is the label distance from the axis, ranging from 0 to 1000 (both inclusive). The default is 100.
        /// </summary>
        public ushort LabelOffset
        {
            get { return iLabelOffset; }
            set
            {
                iLabelOffset = value;
                if (iLabelOffset > 1000) iLabelOffset = 1000;
            }
        }

        /// <summary>
        /// The maximum value on the axis. A null value means it's automatically set. WARNING: This is used for date axes. It's also shared with value axes. If it's set for category axes, chart behaviour is not defined.
        /// </summary>
        public DateTime? MaximumDate
        {
            get
            {
                if (this.MaxAxisValue == null)
                {
                    return null;
                }
                else
                {
                    return SLTool.CalculateDateTimeFromDaysFromEpoch(this.MaxAxisValue.Value, this.Date1904);
                }
            }
            set
            {
                if (value == null)
                {
                    this.MaxAxisValue = null;
                }
                else
                {
                    this.MaxAxisValue = SLTool.CalculateDaysFromEpoch(value.Value, this.Date1904);
                }
            }
        }

        /// <summary>
        /// The minimum value on the axis. A null value means it's automatically set. WARNING: This is used for date axes. It's also shared with value axes. If it's set for category axes, chart behaviour is not defined.
        /// </summary>
        public DateTime? MinimumDate
        {
            get
            {
                if (this.MinAxisValue == null)
                {
                    return null;
                }
                else
                {
                    return SLTool.CalculateDateTimeFromDaysFromEpoch(this.MinAxisValue.Value, this.Date1904);
                }
            }
            set
            {
                if (value == null)
                {
                    this.MinAxisValue = null;
                }
                else
                {
                    this.MinAxisValue = SLTool.CalculateDaysFromEpoch(value.Value, this.Date1904);
                }
            }
        }

        /// <summary>
        /// The maximum value on the axis. A null value means it's automatically set. WARNING: This is used for value axis. It's also shared with date axes. If it's set for category axes, chart behaviour is not defined.
        /// </summary>
        public double? MaximumValue
        {
            get { return this.MaxAxisValue; }
            set { this.MaxAxisValue = value; }
        }

        /// <summary>
        /// The minimum value on the axis. A null value means it's automatically set. WARNING: This is used for value axis. It's also shared with date axes. If it's set for category axes, chart behaviour is not defined.
        /// </summary>
        public double? MinimumValue
        {
            get { return this.MinAxisValue; }
            set { this.MinAxisValue = value; }
        }

        /// <summary>
        /// The major unit on the axis. A null value means it's automatically set. This is for the value axis.
        /// </summary>
        public double? ValueMajorUnit { get; set; }

        /// <summary>
        /// The minor unit on the axis. A null value means it's automatically set. This is for the value axis.
        /// </summary>
        public double? ValueMinorUnit { get; set; }

        /// <summary>
        /// Logarithmic scale of the axis, ranging from 2 to 1000 (both inclusive). A null value means it's not used. This is for the value axis.
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
        /// The base unit for date axes. A null value means it's automatically set.
        /// </summary>
        public C.TimeUnitValues? BaseUnit { get; set; }

        internal int? iMajorUnit;
        internal C.TimeUnitValues vMajorTimeUnit;

        internal int? iMinorUnit;
        internal C.TimeUnitValues vMinorTimeUnit;

        // This is actually for the value axis, but due to the way Excel displays the user interface,
        // this is set on the category/date/value axis settings. I don't understand it either...
        /// <summary>
        /// This sets how the axis crosses regarding the tick marks (or position of the axis). Use Between for "between tick marks", and MidpointCategory for "on tick marks".
        /// </summary>
        public C.CrossBetweenValues CrossBetween { get; set; }

        /// <summary>
        /// Indicates if labels are shown as flat text. If false, then the labels are shown as a hierarchy.
        /// This is used only for category axes. The default is true.
        /// </summary>
        public bool NoMultiLevelLabels { get; set; }

        internal SLTextAxis(List<System.Drawing.Color> ThemeColors, bool Date1904, bool IsStylish, bool ThrowExceptionsIfAny)
            : base(ThemeColors, IsStylish, ThrowExceptionsIfAny)
        {
            this.Date1904 = Date1904;

            this.AxisType = SLAxisType.Category;
            this.AutoLabeled = true;

            this.iTickLabelSkip = 1;
            this.iTickMarkSkip = 1;
            this.iLabelOffset = 100;

            this.ValueMajorUnit = null;
            this.ValueMinorUnit = null;
            this.BuiltInUnitValues = null;
            this.ShowDisplayUnitsLabel = false;

            this.BaseUnit = null;
            this.iMajorUnit = null;
            this.vMajorTimeUnit = C.TimeUnitValues.Days;
            this.iMinorUnit = null;
            this.vMinorTimeUnit = C.TimeUnitValues.Days;

            this.CrossBetween = C.CrossBetweenValues.Between;
            this.LabelAlignment = C.LabelAlignmentValues.Center;

            // it used to be true. I have no idea what this does...
            this.NoMultiLevelLabels = false;

            if (IsStylish)
            {
                this.ShapeProperties.Fill.SetNoFill();
                this.ShapeProperties.Outline.Width = 0.75m;
                this.ShapeProperties.Outline.CapType = A.LineCapValues.Flat;
                this.ShapeProperties.Outline.CompoundLineType = A.CompoundLineValues.Single;
                this.ShapeProperties.Outline.Alignment = A.PenAlignmentValues.Center;
                this.ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                this.ShapeProperties.Outline.JoinType = Drawing.SLLineJoinValues.Round;
            }
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
        }

        // We have SetAsCategoryAxis() and SetAsDateAxis() because
        // we want to keep the option of SetAutomaticAxisType()
        // and that needs examining the chart data to determine the type.
        // I don't feel that's very value-added enough...

        /// <summary>
        /// Set this axis as a category axis. WARNING: This only works if it's a category/date axis. This fails if it's already a value axis.
        /// </summary>
        public void SetAsCategoryAxis()
        {
            if (this.AxisType != SLAxisType.Value)
            {
                this.AxisType = SLAxisType.Category;
                this.AutoLabeled = false;
            }
        }

        /// <summary>
        /// Set this axis as a date axis. WARNING: This only works if it's a category/date axis. This fails if it's already a value axis.
        /// </summary>
        public void SetAsDateAxis()
        {
            if (this.AxisType != SLAxisType.Value)
            {
                this.AxisType = SLAxisType.Date;
                this.AutoLabeled = false;
            }
        }

        /// <summary>
        /// Set the major unit for date axes to be automatic.
        /// </summary>
        public void SetAutomaticDateMajorUnit()
        {
            this.iMajorUnit = null;
            this.vMajorTimeUnit = C.TimeUnitValues.Days;
        }

        /// <summary>
        /// Set the major unit for date axes.
        /// </summary>
        /// <param name="MajorUnit">A positive value. Suggested range is 1 to 999999999 (both inclusive).</param>
        /// <param name="MajorTimeUnit">The time unit.</param>
        public void SetDateMajorUnit(int MajorUnit, C.TimeUnitValues MajorTimeUnit)
        {
            this.iMajorUnit = MajorUnit;
            this.vMajorTimeUnit = MajorTimeUnit;
        }

        /// <summary>
        /// Set the minor unit for date axes to be automatic.
        /// </summary>
        public void SetAutomaticDateMinorUnit()
        {
            this.iMinorUnit = null;
            this.vMinorTimeUnit = C.TimeUnitValues.Days;
        }

        /// <summary>
        /// Set the minor unit for date axes.
        /// </summary>
        /// <param name="MinorUnit">A positive value. Suggested range is 1 to 999999999 (both inclusive).</param>
        /// <param name="MinorTimeUnit">The time unit.</param>
        public void SetDateMinorUnit(int MinorUnit, C.TimeUnitValues MinorTimeUnit)
        {
            this.iMinorUnit = MinorUnit;
            this.vMinorTimeUnit = MinorTimeUnit;
        }

        /// <summary>
        /// Set the display units on the axis. This is for value axis.
        /// </summary>
        /// <param name="BuiltInUnit">Built-in unit types.</param>
        /// <param name="ShowDisplayUnitsLabel">True to show the display units label on the chart. False otherwise.</param>
        public void SetDisplayUnits(C.BuiltInUnitValues BuiltInUnit, bool ShowDisplayUnitsLabel)
        {
            this.BuiltInUnitValues = BuiltInUnit;
            this.ShowDisplayUnitsLabel = ShowDisplayUnitsLabel;
        }

        /// <summary>
        /// Remove the display units on the axis. This is for value axis.
        /// </summary>
        public void RemoveDisplayUnits()
        {
            this.BuiltInUnitValues = null;
            this.ShowDisplayUnitsLabel = false;
        }

        /// <summary>
        /// Set the corresponding value axis to cross this axis at an automatic value.
        /// </summary>
        public void SetAutomaticOtherAxisCrossing()
        {
            this.OtherAxisIsCrosses = true;
            this.OtherAxisCrosses = C.CrossesValues.AutoZero;
            this.OtherAxisCrossesAt = 0;
        }

        /// <summary>
        /// Set the corresponding value axis to cross this axis at a given category number. Suggested range is 1 to 31999 (both inclusive). This is for category axis. WARNING: Internally, this is used for category, date and value axes. Remember to set the axis type.
        /// </summary>
        /// <param name="CategoryNumber">Category number to cross at.</param>
        public void SetOtherAxisCrossing(int CategoryNumber)
        {
            this.OtherAxisIsCrosses = false;
            this.OtherAxisCrosses = C.CrossesValues.AutoZero;
            this.OtherAxisCrossesAt = CategoryNumber;
        }

        /// <summary>
        /// Set the corresponding value axis to cross this axis at a given date. This is for date axis. WARNING: Internally, this is used for category, date and value axes. Remember to set the axis type.
        /// </summary>
        /// <param name="DateToBeCrossed">Date to cross at.</param>
        public void SetOtherAxisCrossing(DateTime DateToBeCrossed)
        {
            this.OtherAxisIsCrosses = false;
            this.OtherAxisCrosses = C.CrossesValues.AutoZero;
            this.OtherAxisCrossesAt = SLTool.CalculateDaysFromEpoch(DateToBeCrossed, this.Date1904);
            // the given date is before the epochs (1900 or 1904).
            // Just set to whatever the current epoch is being used.
            if (this.OtherAxisCrossesAt < 0.0) this.OtherAxisCrossesAt = this.Date1904 ? 0.0 : 1.0;
        }

        /// <summary>
        /// Set the corresponding value axis to cross this axis at a given value. This is for value axis. WARNING: Internally, this is used for category, date and value axes. If it's already a value axis, you can't set the axis type.
        /// </summary>
        /// <param name="CrossingAxisValue">Axis value to cross at.</param>
        public void SetOtherAxisCrossing(double CrossingAxisValue)
        {
            this.OtherAxisIsCrosses = false;
            this.OtherAxisCrosses = C.CrossesValues.AutoZero;
            this.OtherAxisCrossesAt = CrossingAxisValue;
        }

        /// <summary>
        /// Set the corresponding value axis to cross this axis at the maximum value.
        /// </summary>
        public void SetMaximumOtherAxisCrossing()
        {
            this.OtherAxisIsCrosses = true;
            this.OtherAxisCrosses = C.CrossesValues.Maximum;
            this.OtherAxisCrossesAt = 0;
        }

        internal C.CategoryAxis ToCategoryAxis(bool IsStylish)
        {
            C.CategoryAxis ca = new C.CategoryAxis();
            ca.AxisId = new C.AxisId() { Val = this.AxisId };

            ca.Scaling = new C.Scaling();
            ca.Scaling.Orientation = new C.Orientation() { Val = this.Orientation };
            if (this.LogBase != null) ca.Scaling.LogBase = new C.LogBase() { Val = this.LogBase.Value };
            if (this.MaxAxisValue != null) ca.Scaling.MaxAxisValue = new C.MaxAxisValue() { Val = this.MaxAxisValue.Value };
            if (this.MinAxisValue != null) ca.Scaling.MinAxisValue = new C.MinAxisValue() { Val = this.MinAxisValue.Value };

            ca.Delete = new C.Delete() { Val = this.Delete };

            C.AxisPositionValues axpos = this.AxisPosition;
            if (!this.ForceAxisPosition)
            {
                if (this.OtherAxisIsInReverseOrder) axpos = SLChartTool.GetOppositePosition(axpos);
                if (this.OtherAxisCrossedAtMaximum) axpos = SLChartTool.GetOppositePosition(axpos);
            }
            ca.AxisPosition = new C.AxisPosition() { Val = axpos };

            if (this.ShowMajorGridlines)
            {
                ca.MajorGridlines = this.MajorGridlines.ToMajorGridlines(IsStylish);
            }

            if (this.ShowMinorGridlines)
            {
                ca.MinorGridlines = this.MinorGridlines.ToMinorGridlines(IsStylish);
            }

            if (this.ShowTitle)
            {
                ca.Title = this.Title.ToTitle(IsStylish);
            }

            if (this.HasNumberingFormat)
            {
                ca.NumberingFormat = new C.NumberingFormat()
                {
                    FormatCode = this.FormatCode,
                    SourceLinked = this.SourceLinked
                };
            }

            ca.MajorTickMark = new C.MajorTickMark() { Val = this.MajorTickMark };
            ca.MinorTickMark = new C.MinorTickMark() { Val = this.MinorTickMark };
            ca.TickLabelPosition = new C.TickLabelPosition() { Val = this.TickLabelPosition };

            if (this.ShapeProperties.HasShapeProperties) ca.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            if (this.Rotation != null || this.Vertical != null || this.Anchor != null || this.AnchorCenter != null)
            {
                ca.TextProperties = new C.TextProperties();
                ca.TextProperties.BodyProperties = new A.BodyProperties();
                if (this.Rotation != null) ca.TextProperties.BodyProperties.Rotation = (int)(this.Rotation.Value * SLConstants.DegreeToAngleRepresentation);
                if (this.Vertical != null) ca.TextProperties.BodyProperties.Vertical = this.Vertical.Value;
                if (this.Anchor != null) ca.TextProperties.BodyProperties.Anchor = this.Anchor.Value;
                if (this.AnchorCenter != null) ca.TextProperties.BodyProperties.AnchorCenter = this.AnchorCenter.Value;

                ca.TextProperties.ListStyle = new A.ListStyle();

                A.Paragraph para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();
                para.ParagraphProperties.Append(new A.DefaultRunProperties());
                ca.TextProperties.Append(para);
            }
            else if (IsStylish)
            {
                ca.TextProperties = new C.TextProperties();
                ca.TextProperties.BodyProperties = new A.BodyProperties()
                {
                    Rotation = -60000000,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                ca.TextProperties.ListStyle = new A.ListStyle();

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

                ca.TextProperties.Append(para);
            }

            ca.CrossingAxis = new C.CrossingAxis() { Val = this.CrossingAxis };

            if (this.IsCrosses != null)
            {
                if (this.IsCrosses.Value)
                {
                    ca.Append(new C.Crosses() { Val = this.Crosses });
                }
                else
                {
                    ca.Append(new C.CrossesAt() { Val = this.CrossesAt });
                }
            }

            ca.Append(new C.AutoLabeled() { Val = this.AutoLabeled });
            ca.Append(new C.LabelAlignment() { Val = this.LabelAlignment });
            ca.Append(new C.LabelOffset() { Val = this.LabelOffset });

            if (this.iTickLabelSkip > 1) ca.Append(new C.TickLabelSkip() { Val = this.TickLabelSkip });
            if (this.iTickMarkSkip > 1) ca.Append(new C.TickMarkSkip() { Val = this.TickMarkSkip });

            ca.Append(new C.NoMultiLevelLabels() { Val = this.NoMultiLevelLabels });

            return ca;
        }

        internal C.DateAxis ToDateAxis(bool IsStylish)
        {
            C.DateAxis da = new C.DateAxis();
            da.AxisId = new C.AxisId() { Val = this.AxisId };

            da.Scaling = new C.Scaling();
            da.Scaling.Orientation = new C.Orientation() { Val = this.Orientation };
            if (this.LogBase != null) da.Scaling.LogBase = new C.LogBase() { Val = this.LogBase.Value };
            if (this.MaxAxisValue != null) da.Scaling.MaxAxisValue = new C.MaxAxisValue() { Val = this.MaxAxisValue.Value };
            if (this.MinAxisValue != null) da.Scaling.MinAxisValue = new C.MinAxisValue() { Val = this.MinAxisValue.Value };

            da.Delete = new C.Delete() { Val = this.Delete };

            C.AxisPositionValues axpos = this.AxisPosition;
            if (!this.ForceAxisPosition)
            {
                if (this.OtherAxisIsInReverseOrder) axpos = SLChartTool.GetOppositePosition(axpos);
                if (this.OtherAxisCrossedAtMaximum) axpos = SLChartTool.GetOppositePosition(axpos);
            }
            da.AxisPosition = new C.AxisPosition() { Val = axpos };

            if (this.ShowMajorGridlines)
            {
                da.MajorGridlines = this.MajorGridlines.ToMajorGridlines(IsStylish);
            }

            if (this.ShowMinorGridlines)
            {
                da.MinorGridlines = this.MinorGridlines.ToMinorGridlines(IsStylish);
            }

            if (this.ShowTitle)
            {
                da.Title = this.Title.ToTitle(IsStylish);
            }

            if (this.HasNumberingFormat)
            {
                da.NumberingFormat = new C.NumberingFormat()
                {
                    FormatCode = this.FormatCode,
                    SourceLinked = this.SourceLinked
                };
            }

            da.MajorTickMark = new C.MajorTickMark() { Val = this.MajorTickMark };
            da.MinorTickMark = new C.MinorTickMark() { Val = this.MinorTickMark };
            da.TickLabelPosition = new C.TickLabelPosition() { Val = this.TickLabelPosition };

            if (this.ShapeProperties.HasShapeProperties) da.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            if (this.Rotation != null || this.Vertical != null || this.Anchor != null || this.AnchorCenter != null)
            {
                da.TextProperties = new C.TextProperties();
                da.TextProperties.BodyProperties = new A.BodyProperties();
                if (this.Rotation != null) da.TextProperties.BodyProperties.Rotation = (int)(this.Rotation.Value * SLConstants.DegreeToAngleRepresentation);
                if (this.Vertical != null) da.TextProperties.BodyProperties.Vertical = this.Vertical.Value;
                if (this.Anchor != null) da.TextProperties.BodyProperties.Anchor = this.Anchor.Value;
                if (this.AnchorCenter != null) da.TextProperties.BodyProperties.AnchorCenter = this.AnchorCenter.Value;

                da.TextProperties.ListStyle = new A.ListStyle();

                A.Paragraph para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();
                para.ParagraphProperties.Append(new A.DefaultRunProperties());
                da.TextProperties.Append(para);
            }
            else if (IsStylish)
            {
                da.TextProperties = new C.TextProperties();
                da.TextProperties.BodyProperties = new A.BodyProperties()
                {
                    Rotation = -60000000,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                da.TextProperties.ListStyle = new A.ListStyle();

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

                da.TextProperties.Append(para);
            }

            da.CrossingAxis = new C.CrossingAxis() { Val = this.CrossingAxis };

            if (this.IsCrosses != null)
            {
                if (this.IsCrosses.Value)
                {
                    da.Append(new C.Crosses() { Val = this.Crosses });
                }
                else
                {
                    da.Append(new C.CrossesAt() { Val = this.CrossesAt });
                }
            }

            da.Append(new C.AutoLabeled() { Val = this.AutoLabeled });
            da.Append(new C.LabelOffset() { Val = this.LabelOffset });

            if (this.BaseUnit != null) da.Append(new C.BaseTimeUnit() { Val = this.BaseUnit.Value });

            if (this.iMajorUnit != null)
            {
                da.Append(new C.MajorUnit() { Val = this.iMajorUnit.Value });
                da.Append(new C.MajorTimeUnit() { Val = this.vMajorTimeUnit });
            }

            if (this.iMinorUnit != null)
            {
                da.Append(new C.MinorUnit() { Val = this.iMinorUnit.Value });
                da.Append(new C.MinorTimeUnit() { Val = this.vMinorTimeUnit });
            }

            return da;
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
            if (this.ValueMajorUnit != null) va.Append(new C.MajorUnit() { Val = this.ValueMajorUnit.Value });
            if (this.ValueMinorUnit != null) va.Append(new C.MinorUnit() { Val = this.ValueMinorUnit.Value });

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

        internal SLTextAxis Clone()
        {
            SLTextAxis ta = new SLTextAxis(this.ShapeProperties.listThemeColors, this.Date1904, false, this.ShapeProperties.ThrowExceptionsIfAny);
            ta.Rotation = this.Rotation;
            ta.Vertical = this.Vertical;
            ta.Anchor = this.Anchor;
            ta.AnchorCenter = this.AnchorCenter;
            ta.AxisId = this.AxisId;
            ta.fLogBase = this.fLogBase;
            ta.Orientation = this.Orientation;
            ta.MaxAxisValue = this.MaxAxisValue;
            ta.MinAxisValue = this.MinAxisValue;
            ta.OtherAxisIsInReverseOrder = this.OtherAxisIsInReverseOrder;
            ta.OtherAxisCrossedAtMaximum = this.OtherAxisCrossedAtMaximum;
            ta.Delete = this.Delete;
            ta.ForceAxisPosition = this.ForceAxisPosition;
            ta.AxisPosition = this.AxisPosition;
            ta.ShowMajorGridlines = this.ShowMajorGridlines;
            ta.MajorGridlines = this.MajorGridlines.Clone();
            ta.ShowMinorGridlines = this.ShowMinorGridlines;
            ta.MinorGridlines = this.MinorGridlines.Clone();
            ta.ShowTitle = this.ShowTitle;
            ta.Title = this.Title.Clone();
            ta.HasNumberingFormat = this.HasNumberingFormat;
            ta.sFormatCode = this.sFormatCode;
            ta.bSourceLinked = this.bSourceLinked;
            ta.MajorTickMark = this.MajorTickMark;
            ta.MinorTickMark = this.MinorTickMark;
            ta.TickLabelPosition = this.TickLabelPosition;
            ta.ShapeProperties = this.ShapeProperties.Clone();
            ta.CrossingAxis = this.CrossingAxis;
            ta.IsCrosses = this.IsCrosses;
            ta.Crosses = this.Crosses;
            ta.CrossesAt = this.CrossesAt;
            ta.OtherAxisIsCrosses = this.OtherAxisIsCrosses;
            ta.OtherAxisCrosses = this.OtherAxisCrosses;
            ta.OtherAxisCrossesAt = this.OtherAxisCrossesAt;

            ta.Date1904 = this.Date1904;
            ta.AxisType = this.AxisType;
            ta.AutoLabeled = this.AutoLabeled;
            ta.iTickLabelSkip = this.iTickLabelSkip;
            ta.iTickMarkSkip = this.iTickMarkSkip;
            ta.LabelAlignment = this.LabelAlignment;
            ta.iLabelOffset = this.iLabelOffset;
            ta.ValueMajorUnit = this.ValueMajorUnit;
            ta.ValueMinorUnit = this.ValueMinorUnit;
            ta.BuiltInUnitValues = this.BuiltInUnitValues;
            ta.ShowDisplayUnitsLabel = this.ShowDisplayUnitsLabel;
            ta.BaseUnit = this.BaseUnit;
            ta.iMajorUnit = this.iMajorUnit;
            ta.vMajorTimeUnit = this.vMajorTimeUnit;
            ta.iMinorUnit = this.iMinorUnit;
            ta.vMinorTimeUnit = this.vMinorTimeUnit;
            ta.CrossBetween = this.CrossBetween;
            ta.NoMultiLevelLabels = this.NoMultiLevelLabels;

            return ta;
        }
    }
}
