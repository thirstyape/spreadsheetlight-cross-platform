using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting data point options for charts.
    /// </summary>
    public class SLDataPointOptions
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

        // pictureoptions?

        internal SLDataPointOptions(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
            this.InvertIfNegative = null;
            this.Marker = new SLMarker(ThemeColors, ThrowExceptionsIfAny);
            this.iExplosion = null;
            this.bBubble3D = null;
        }

        internal C.DataPoint ToDataPoint(int index, bool IsStylish)
        {
            C.DataPoint pt = new C.DataPoint();

            pt.Index = new C.Index() { Val = (uint)index };

            if (this.Marker.HasMarker) pt.Marker = this.Marker.ToMarker(IsStylish);

            if (this.bBubble3D != null) pt.Bubble3D = new C.Bubble3D() { Val = this.bBubble3D.Value };

            if (this.iExplosion != null) pt.Explosion = new C.Explosion() { Val = this.iExplosion.Value };

            if (this.ShapeProperties.HasShapeProperties) pt.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            return pt;
        }

        internal SLDataPointOptions Clone()
        {
            SLDataPointOptions dpo = new SLDataPointOptions(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
            dpo.ShapeProperties = this.ShapeProperties.Clone();
            dpo.InvertIfNegative = this.InvertIfNegative;
            dpo.Marker = this.Marker.Clone();
            dpo.iExplosion = this.iExplosion;
            dpo.bBubble3D = this.bBubble3D;

            return dpo;
        }
    }
}
