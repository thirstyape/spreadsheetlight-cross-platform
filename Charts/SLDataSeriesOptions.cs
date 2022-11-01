using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Data series customization options. Note that all supported chart data series properties are available, but only the relevant properties (to chart type) will be used.
    /// </summary>
    public class SLDataSeriesOptions
    {
        internal SLA.SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Fill properties.
        /// </summary>
        public SLA.SLFill Fill { get { return this.ShapeProperties.Fill; } }

        /// <summary>
        /// Border/Line properties.
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

        // internally, the default is actually true in Open XML, but when null it's false.
        // The Open XML docs state it's supposed to be true when the tag is missing. I don't know...
        /// <summary>
        /// Invert colors if negative. If null, the effective default is used (false). This is for bar charts, column charts and bubble charts.
        /// </summary>
        public bool? InvertIfNegative { get; set; }

        /// <summary>
        /// Marker properties. This is for line charts, radar charts and scatter charts.
        /// </summary>
        public SLMarker Marker { get; set; }

        // "default" is 25%, range of 0% to 400%
        // but we're not enforcing the range
        internal uint? iExplosion;
        /// <summary>
        /// The explosion distance from the center of the pie in percentage. It is suggested to keep the range between 0% and 400%.
        /// </summary>
        public uint Explosion
        {
            get { return iExplosion ?? 0; }
            set { iExplosion = value; }
        }

        internal bool? bBubble3D;
        internal bool Bubble3D
        {
            get { return bBubble3D ?? true; }
            set { bBubble3D = value; }
        }

        /// <summary>
        /// Whether the line connecting data points use C splines (instead of straight lines). This is for line charts and scatter charts.
        /// </summary>
        public bool Smooth { get; set; }

        internal C.ShapeValues? vShape;
        /// <summary>
        /// The shape of data series for 3D bar and column charts.
        /// </summary>
        public C.ShapeValues Shape
        {
            get { return vShape ?? C.ShapeValues.Box; }
            set { vShape = value; }
        }

        /// <summary>
        /// Initializes an instance of SLDataSeriesOptions. It is recommended to use SLChart.GetDataSeriesOptions().
        /// </summary>
        public SLDataSeriesOptions()
        {
            this.Initialize(new List<System.Drawing.Color>(), false);
        }

        internal SLDataSeriesOptions(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            this.Initialize(ThemeColors, ThrowExceptionsIfAny);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
            this.InvertIfNegative = null;
            this.Marker = new SLMarker(ThemeColors, ThrowExceptionsIfAny);
            this.iExplosion = null;
            this.bBubble3D = null;
            this.Smooth = false;
            this.vShape = null;
        }

        internal SLDataSeriesOptions Clone()
        {
            SLDataSeriesOptions dso = new SLDataSeriesOptions(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
            dso.ShapeProperties = this.ShapeProperties.Clone();
            dso.InvertIfNegative = this.InvertIfNegative;
            dso.Marker = this.Marker.Clone();
            dso.iExplosion = this.iExplosion;
            dso.bBubble3D = this.bBubble3D;
            dso.Smooth = this.Smooth;
            dso.vShape = this.vShape;

            return dso;
        }
    }
}
