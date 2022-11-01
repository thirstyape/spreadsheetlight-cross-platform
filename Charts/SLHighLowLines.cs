using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for high-low lines.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.HighLowLines class.
    /// </summary>
    public class SLHighLowLines
    {
        internal SLA.SLShapeProperties ShapeProperties { get; set; }

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

        internal SLHighLowLines(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
            if (IsStylish)
            {
                this.ShapeProperties.Outline.Width = 0.75m;
                this.ShapeProperties.Outline.CapType = A.LineCapValues.Flat;
                this.ShapeProperties.Outline.CompoundLineType = A.CompoundLineValues.Single;
                this.ShapeProperties.Outline.Alignment = A.PenAlignmentValues.Center;
                this.ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.25m, 0);
                this.ShapeProperties.Outline.JoinType = SLA.SLLineJoinValues.Round;
            }
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
        }

        internal C.HighLowLines ToHighLowLines(bool IsStylish)
        {
            C.HighLowLines hll = new C.HighLowLines();

            if (this.ShapeProperties.HasShapeProperties) hll.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            return hll;
        }

        internal SLHighLowLines Clone()
        {
            SLHighLowLines hll = new SLHighLowLines(this.ShapeProperties.listThemeColors, false, this.ShapeProperties.ThrowExceptionsIfAny);
            hll.ShapeProperties = this.ShapeProperties.Clone();

            return hll;
        }
    }
}
