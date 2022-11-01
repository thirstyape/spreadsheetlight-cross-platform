using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    public class SLTrendlineLabel
    {
        internal SLLayout Layout { get; set; }

        /// <summary>
        /// Format code for the trendline label.
        /// </summary>
        public string FormatCode { get; set; }

        /// <summary>
        /// Whether the format code is linked to the data source.
        /// </summary>
        public bool SourceLinked { get; set; }

        internal SLA.SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Fill properties.
        /// </summary>
        public SLA.SLFill Fill { get { return this.ShapeProperties.Fill; } }

        /// <summary>
        /// Border properties.
        /// </summary>
        public SLA.SLLinePropertiesType Border { get { return this.ShapeProperties.Outline; } }

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

        // TODO: ignoring rich text content
        // TODO: ignoring text alignment properties for now
        // TODO: ignoring text body properties for now

        internal SLTrendlineLabel(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            this.Layout = new SLLayout();
            this.FormatCode = "General";
            this.SourceLinked = false;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
        }

        internal C.TrendlineLabel ToTrendlineLabel(bool IsStylish)
        {
            C.TrendlineLabel tll = new C.TrendlineLabel();
            tll.Layout = this.Layout.ToLayout();
            tll.NumberingFormat = new C.NumberingFormat();
            tll.NumberingFormat.FormatCode = this.FormatCode;
            tll.NumberingFormat.SourceLinked = this.SourceLinked;

            if (this.ShapeProperties.HasShapeProperties)
            {
                tll.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            return tll;
        }
    }
}
