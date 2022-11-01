using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting the floor of 3D charts.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.Floor class.
    /// </summary>
    public class SLFloor
    {
        // From the Open XML SDK documentation:
        // "This element specifies the thickness of the walls or floor as a percentage of the largest dimension of the plot volume."
        // I have no idea what that means... and Excel doesn't allow the user to set this. Hmm...
        internal byte Thickness { get; set; }

        internal SLA.SLShapeProperties ShapeProperties;

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

        /// <summary>
        /// 3D rotation properties.
        /// </summary>
        public SLA.SLRotation3D Rotation3D { get { return this.ShapeProperties.Rotation3D; } }

        internal SLFloor(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.Thickness = 0;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
            if (IsStylish)
            {
                this.ShapeProperties.Fill.SetNoFill();
                this.ShapeProperties.Outline.Width = 0.75m;
                this.ShapeProperties.Outline.CapType = A.LineCapValues.Flat;
                this.ShapeProperties.Outline.CompoundLineType = A.CompoundLineValues.Single;
                this.ShapeProperties.Outline.Alignment = A.PenAlignmentValues.Center;
                this.ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                this.ShapeProperties.Outline.JoinType = SLA.SLLineJoinValues.Round;
                this.ShapeProperties.Format3D.ContourWidth = 0.75m;
                this.ShapeProperties.Format3D.clrContourColor.SetColor(A.SchemeColorValues.Text1, 0.85m, 0);
                this.ShapeProperties.Format3D.HasContourColor = true;
            }
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
        }

        internal SLFloor Clone()
        {
            SLFloor f = new SLFloor(this.ShapeProperties.listThemeColors, false, this.ShapeProperties.ThrowExceptionsIfAny);
            f.Thickness = this.Thickness;
            f.ShapeProperties = this.ShapeProperties.Clone();

            return f;
        }
    }
}
