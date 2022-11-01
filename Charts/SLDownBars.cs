using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for down bars.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.DownBars class.
    /// </summary>
    public class SLDownBars
    {
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

        internal SLDownBars(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
            if (IsStylish)
            {
                this.ShapeProperties.Fill.SetSolidFill(A.SchemeColorValues.Dark1, 0.35m, 0);
                this.ShapeProperties.Outline.Width = 0.75m;
                this.ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.35m, 0);
            }
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
        }

        internal C.DownBars ToDownBars(bool IsStylish)
        {
            C.DownBars db = new C.DownBars();

            if (this.ShapeProperties.HasShapeProperties) db.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            return db;
        }

        internal SLDownBars Clone()
        {
            SLDownBars db = new SLDownBars(this.ShapeProperties.listThemeColors, false, this.ShapeProperties.ThrowExceptionsIfAny);
            db.ShapeProperties = this.ShapeProperties.Clone();

            return db;
        }
    }
}
