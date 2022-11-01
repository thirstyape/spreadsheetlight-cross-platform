using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// For CategoryAxis, ValueAxis, SeriesAxis and DateAxis from namespace DocumentFormat.OpenXml.Drawing.Charts.
    /// </summary>
    public abstract class EGAxShared : SLChartAlignment
    {
        internal uint AxisId { get; set; }

        // Scaling
        internal double? fLogBase;
        internal double? LogBase
        {
            get { return fLogBase; }
            set
            {
                fLogBase = value;
                if (value != null)
                {
                    if (value.Value < 2.0) fLogBase = 2.0;
                    if (value.Value > 1000.0) fLogBase = 1000.0;
                }
            }
        }

        internal C.OrientationValues Orientation { get; set; }
        internal double? MaxAxisValue { get; set; }
        internal double? MinAxisValue { get; set; }

        /// <summary>
        /// Display axis values in reverse order.
        /// </summary>
        public bool InReverseOrder
        {
            get { return this.Orientation != C.OrientationValues.MinMax; }
            set
            {
                if (value)
                {
                    this.Orientation = C.OrientationValues.MaxMin;
                    return;
                }
                this.Orientation = C.OrientationValues.MinMax;
            }
        }

        internal bool OtherAxisIsInReverseOrder;
        internal bool OtherAxisCrossedAtMaximum;

        internal bool Delete { get; set; }

        internal bool ForceAxisPosition { get; set; }
        internal C.AxisPositionValues AxisPosition { get; set; }

        /// <summary>
        /// Whether major gridlines are shown.
        /// </summary>
        public bool ShowMajorGridlines { get; set; }

        /// <summary>
        /// Major gridlines properties.
        /// </summary>
        public SLMajorGridlines MajorGridlines { get; set; }

        /// <summary>
        /// Whether minor gridlines are shown.
        /// </summary>
        public bool ShowMinorGridlines { get; set; }

        /// <summary>
        /// Minor gridlines properties.
        /// </summary>
        public SLMinorGridlines MinorGridlines { get; set; }

        /// <summary>
        /// Whether the axis title is shown.
        /// </summary>
        public bool ShowTitle { get; set; }

        /// <summary>
        /// Axis title properties.
        /// </summary>
        public SLTitle Title { get; set; }

        // This is C.NumberingFormat
        internal bool HasNumberingFormat;

        internal string sFormatCode;
        /// <summary>
        /// Format code for the axis. If you set a custom format code, you might also want to set SourceLinked to false.
        /// </summary>
        public string FormatCode
        {
            get { return sFormatCode; }
            set
            {
                sFormatCode = value;
                HasNumberingFormat = true;
            }
        }

        internal bool bSourceLinked;
        /// <summary>
        /// Whether the format code is linked to the data source.
        /// </summary>
        public bool SourceLinked
        {
            get { return bSourceLinked; }
            set
            {
                bSourceLinked = value;
                HasNumberingFormat = true;
            }
        }

        /// <summary>
        /// Major tick mark type.
        /// </summary>
        public C.TickMarkValues MajorTickMark { get; set; }

        /// <summary>
        /// Minor tick mark type.
        /// </summary>
        public C.TickMarkValues MinorTickMark { get; set; }

        /// <summary>
        /// Position of axis labels.
        /// </summary>
        public C.TickLabelPositionValues TickLabelPosition { get; set; }

        // C.ChartShapeProperties
        internal SLA.SLShapeProperties ShapeProperties;

        /// <summary>
        /// Fill properties.
        /// </summary>
        public SLA.SLFill Fill { get { return this.ShapeProperties.Fill; } }

        /// <summary>
        /// Line properties.
        /// </summary>
        public SLA.SLLinePropertiesType Line { get { return this.ShapeProperties.Outline; } }

        /// <summary>
        /// Shadow properties.
        /// </summary>
        public SLA.SLShadowEffect Shadow { get { return this.ShapeProperties.EffectList.Shadow; } }

        /// <summary>
        /// Glow properties.
        /// </summary>
        public SLA.SLGlow Glow { get { return this.ShapeProperties.EffectList.Glow; } }

        /// <summary>
        /// Soft edge properties.
        /// </summary>
        public SLA.SLSoftEdge SoftEdge { get { return this.ShapeProperties.EffectList.SoftEdge; } }

        /// <summary>
        /// 3D format properties.
        /// </summary>
        public SLA.SLFormat3D Format3D { get { return this.ShapeProperties.Format3D; } }

        // C.TextProperties

        internal uint CrossingAxis { get; set; }

        internal bool? IsCrosses;
        internal C.CrossesValues Crosses { get; set; }
        internal double CrossesAt { get; set; }

        // The Excel UI sets cross values of that axis for the *other* axis.
        // Weird... Meaning if you set the cross value of the category axis (at least
        // on the UI), you're actually setting the cross value of the value axis.
        // Why Excel didn't set them on the actual axis is beyond me...
        // Maybe on the UI it made sense to do so.
        internal bool? OtherAxisIsCrosses;
        internal C.CrossesValues OtherAxisCrosses { get; set; }
        internal double OtherAxisCrossesAt { get; set; }

        internal EGAxShared(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.AxisId = 0;
            this.LogBase = null;
            this.Orientation = C.OrientationValues.MinMax;

            this.OtherAxisIsInReverseOrder = false;
            this.OtherAxisCrossedAtMaximum = false;

            this.MaxAxisValue = null;
            this.MinAxisValue = null;
            this.Delete = false;
            this.ForceAxisPosition = false;
            this.AxisPosition = C.AxisPositionValues.Bottom;

            this.ShowMajorGridlines = false;
            this.MajorGridlines = new SLMajorGridlines(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.ShowMinorGridlines = false;
            this.MinorGridlines = new SLMinorGridlines(ThemeColors, IsStylish, ThrowExceptionsIfAny);

            this.ShowTitle = false;
            this.Title = new SLTitle(ThemeColors, IsStylish, ThrowExceptionsIfAny);

            this.sFormatCode = SLConstants.NumberFormatGeneral;
            this.bSourceLinked = true;
            this.HasNumberingFormat = false;

            this.MajorTickMark = C.TickMarkValues.Outside;
            this.MinorTickMark = C.TickMarkValues.None;
            this.TickLabelPosition = C.TickLabelPositionValues.NextTo; // default

            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);

            this.CrossingAxis = 0;
            this.IsCrosses = null;
            this.Crosses = C.CrossesValues.AutoZero;
            this.CrossesAt = 0.0;

            this.OtherAxisIsCrosses = null;
            this.OtherAxisCrosses = C.CrossesValues.AutoZero;
            this.OtherAxisCrossesAt = 0.0;
        }
    }
}
