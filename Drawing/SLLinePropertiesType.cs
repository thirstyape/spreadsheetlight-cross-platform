using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing
{
    /// <summary>
    /// Encapsulates properties and methods for setting line or border settings.
    /// This simulates the DocumentFormat.OpenXml.Drawing.LinePropertiesType class.
    /// </summary>
    public class SLLinePropertiesType
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal bool ThrowExceptionsIfAny { get; set; }

        internal bool HasLine
        {
            get { return UseNoLine || UseSolidLine || UseGradientLine || HasWidth || HasCapType || HasCompoundLineType || HasDashType || HasJoinType; }
        }

        private bool bUseNoLine = false;
        internal bool UseNoLine
        {
            get { return bUseNoLine; }
            set
            {
                bUseNoLine = value;
                if (value)
                {
                    bUseNoLine = true;
                    bUseSolidLine = false;
                    bUseGradientLine = false;
                }
            }
        }

        private bool bUseSolidLine = false;
        internal bool UseSolidLine
        {
            get { return bUseSolidLine; }
            set
            {
                bUseSolidLine = value;
                if (value)
                {
                    bUseNoLine = false;
                    bUseSolidLine = true;
                    bUseGradientLine = false;
                }
            }
        }
        internal SLColorTransform SolidColor { get; set; }

        private bool bUseGradientLine = false;
        internal bool UseGradientLine
        {
            get { return bUseGradientLine; }
            set
            {
                bUseGradientLine = value;
                if (value)
                {
                    bUseNoLine = false;
                    bUseSolidLine = false;
                    bUseGradientLine = true;
                }
            }
        }
        internal SLGradientFill GradientColor { get; set; }

        internal bool HasDashType = false;
        private A.PresetLineDashValues vDashType;
        /// <summary>
        /// The dash type.
        /// </summary>
        public A.PresetLineDashValues DashType
        {
            get { return vDashType; }
            set
            {
                this.HasDashType = true;
                vDashType = value;
            }
        }

        internal bool HasJoinType = false;
        private SLLineJoinValues vJoinType;
        /// <summary>
        /// The join type.
        /// </summary>
        public SLLineJoinValues JoinType
        {
            get { return vJoinType; }
            set
            {
                this.HasJoinType = true;
                vJoinType = value;
            }
        }

        internal A.LineEndValues? HeadEndType { get; set; }
        internal SLLineSizeValues HeadEndSize { get; set; }
        internal A.LineEndValues? TailEndType { get; set; }
        internal SLLineSizeValues TailEndSize { get; set; }

        internal bool HasWidth = false;
        private decimal decWidth;
        /// <summary>
        /// Width between 0 pt and 1584 pt. Accurate to 1/12700 of a point.
        /// </summary>
        public decimal Width
        {
            get { return decWidth; }
            set
            {
                this.HasWidth = true;
                decWidth = value;
                if (decWidth < 0m) decWidth = 0m;
                if (decWidth > 1584m) decWidth = 1584m;
            }
        }

        internal bool HasCapType = false;
        private A.LineCapValues vCapType;
        /// <summary>
        /// The cap type.
        /// </summary>
        public A.LineCapValues CapType
        {
            get { return vCapType; }
            set
            {
                this.HasCapType = true;
                vCapType = value;
            }
        }

        internal bool HasCompoundLineType = false;
        private A.CompoundLineValues vCompoundLineType;
        /// <summary>
        /// The compound type.
        /// </summary>
        public A.CompoundLineValues CompoundLineType
        {
            get { return vCompoundLineType; }
            set
            {
                this.HasCompoundLineType = true;
                vCompoundLineType = value;
            }
        }

        /// <summary>
        /// The alignment.
        /// </summary>
        public A.PenAlignmentValues? Alignment { get; set; }

        internal SLLinePropertiesType(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }

            this.ThrowExceptionsIfAny = ThrowExceptionsIfAny;

            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.bUseNoLine = false;
            this.bUseSolidLine = false;
            this.SolidColor = new SLColorTransform(this.listThemeColors);
            this.bUseGradientLine = false;
            this.GradientColor = new SLGradientFill(this.listThemeColors, this.ThrowExceptionsIfAny);

            this.decWidth = 0m;
            this.HasWidth = false;
            this.vCompoundLineType = A.CompoundLineValues.Single;
            this.HasCompoundLineType = false;
            this.vDashType = A.PresetLineDashValues.Solid;
            this.HasDashType = false;
            this.vCapType = A.LineCapValues.Square;
            this.HasCapType = false;
            this.vJoinType = SLLineJoinValues.Round;
            this.HasJoinType = false;

            this.HeadEndType = null;
            this.HeadEndSize = SLLineSizeValues.Size1;
            this.TailEndType = null;
            this.TailEndSize = SLLineSizeValues.Size1;

            this.Alignment = null;
        }

        /// <summary>
        /// Set color to be automatic.
        /// </summary>
        public void SetAutomaticColor()
        {
            this.bUseNoLine = false;
            this.bUseSolidLine = false;
            this.bUseGradientLine = false;
        }

        /// <summary>
        /// Set no line.
        /// </summary>
        public void SetNoLine()
        {
            this.UseNoLine = true;
        }

        /// <summary>
        /// Set a solid line given a color for the line and the transparency of the color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetSolidLine(System.Drawing.Color Color, decimal Transparency)
        {
            this.UseSolidLine = true;
            this.SolidColor.SetColor(Color, Transparency);
        }

        /// <summary>
        /// Set a solid line given a color for the line and the transparency of the color.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetSolidLine(SLThemeColorIndexValues Color, double Tint, decimal Transparency)
        {
            this.UseSolidLine = true;
            this.SolidColor.SetColor(Color, Tint, Transparency);
        }

        internal void SetSolidLine(A.SchemeColorValues Color, decimal Tint, decimal Transparency)
        {
            this.UseSolidLine = true;
            this.SolidColor.SetColor(Color, Tint, Transparency);
        }

        /// <summary>
        /// Set a linear gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Angle">The interpolation angle ranging from 0 degrees to 359.9 degrees. 0 degrees mean from left to right, 90 degrees mean from top to bottom, 180 degrees mean from right to left and 270 degrees mean from bottom to top. Accurate to 1/60000 of a degree.</param>
        public void SetLinearGradient(SLGradientPresetValues Preset, decimal Angle)
        {
            this.UseGradientLine = true;
            this.GradientColor.SetLinearGradient(Preset, Angle);
        }

        /// <summary>
        /// Set a radial gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Direction">The radial gradient direction.</param>
        public void SetRadialGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            this.UseGradientLine = true;
            this.GradientColor.SetRadialGradient(Preset, Direction);
        }

        /// <summary>
        /// Set a rectangular gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Direction">The rectangular gradient direction.</param>
        public void SetRectangularGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            this.UseGradientLine = true;
            this.GradientColor.SetRectangularGradient(Preset, Direction);
        }

        /// <summary>
        /// Set a path gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        public void SetPathGradient(SLGradientPresetValues Preset)
        {
            this.UseGradientLine = true;
            this.GradientColor.SetPathGradient(Preset);
        }

        /// <summary>
        /// Append a gradient stop given a color, the color's transparency and the position of gradient stop.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Position">The position in percentage ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void AppendGradientStop(System.Drawing.Color Color, decimal Transparency, decimal Position)
        {
            this.GradientColor.AppendGradientStop(Color, Transparency, Position);
        }

        /// <summary>
        /// Append a gradient stop given a color, the color's transparency and the position of gradient stop.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Position">The position in percentage ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void AppendGradientStop(SLThemeColorIndexValues Color, double Tint, decimal Transparency, decimal Position)
        {
            this.GradientColor.AppendGradientStop(Color, Tint, Transparency, Position);
        }

        /// <summary>
        /// Clear all gradient stops.
        /// </summary>
        public void ClearGradientStops()
        {
            this.GradientColor.ClearGradientStops();
        }

        /// <summary>
        /// Set line arrow head settings. This only makes sense for lines and not border lines.
        /// </summary>
        /// <param name="HeadType">The arrow head type.</param>
        /// <param name="HeadSize">The arrow head size.</param>
        public void SetArrowHead(A.LineEndValues HeadType, SLLineSizeValues HeadSize)
        {
            this.HeadEndType = HeadType;
            this.HeadEndSize = HeadSize;
        }

        /// <summary>
        /// Set line arrow tail settings. This only makes sense for lines and not border lines.
        /// </summary>
        /// <param name="TailType">The arrow tail type.</param>
        /// <param name="TailSize">The arrow tail size.</param>
        public void SetArrowTail(A.LineEndValues TailType, SLLineSizeValues TailSize)
        {
            this.TailEndType = TailType;
            this.TailEndSize = TailSize;
        }

        internal A.Outline ToOutline()
        {
            A.Outline ol = new A.Outline();
            if (this.UseNoLine) ol.Append(new A.NoFill());
            if (this.UseSolidLine)
            {
                if (this.SolidColor.IsRgbColorModelHex)
                {
                    ol.Append(new A.SolidFill() { RgbColorModelHex = this.SolidColor.ToRgbColorModelHex() });
                }
                else
                {
                    ol.Append(new A.SolidFill() { SchemeColor = this.SolidColor.ToSchemeColor() });
                }
            }
            if (this.UseGradientLine)
            {
                ol.Append(this.GradientColor.ToGradientFill());
            }

            if (this.HasDashType) ol.Append(new A.PresetDash() { Val = this.DashType });

            if (this.HasJoinType)
            {
                switch (this.JoinType)
                {
                    case SLLineJoinValues.Round:
                        ol.Append(new A.Round());
                        break;
                    case SLLineJoinValues.Bevel:
                        ol.Append(new A.Bevel());
                        break;
                    case SLLineJoinValues.Miter:
                        // 800000 was the default Excel gave
                        ol.Append(new A.Miter() { Limit = 800000 });
                        break;
                }
            }

            if (this.HeadEndType != null) ol.Append(this.GetHeadEnd());
            if (this.TailEndType != null) ol.Append(this.GetTailEnd());

            if (this.HasWidth) ol.Width = Convert.ToInt32(this.Width * SLConstants.PointToEMU);
            if (this.HasCapType) ol.CapType = this.CapType;
            if (this.HasCompoundLineType) ol.CompoundLineType = this.CompoundLineType;
            if (this.Alignment != null) ol.Alignment = this.Alignment.Value;

            return ol;
        }

        private A.HeadEnd GetHeadEnd()
        {
            A.HeadEnd he = new A.HeadEnd() { Type = this.HeadEndType.Value };
            switch (this.HeadEndSize)
            {
                case SLLineSizeValues.Size1:
                    he.Width = A.LineEndWidthValues.Small;
                    he.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size2:
                    he.Width = A.LineEndWidthValues.Small;
                    he.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size3:
                    he.Width = A.LineEndWidthValues.Small;
                    he.Length = A.LineEndLengthValues.Large;
                    break;
                case SLLineSizeValues.Size4:
                    he.Width = A.LineEndWidthValues.Medium;
                    he.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size5:
                    he.Width = A.LineEndWidthValues.Medium;
                    he.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size6:
                    he.Width = A.LineEndWidthValues.Medium;
                    he.Length = A.LineEndLengthValues.Large;
                    break;
                case SLLineSizeValues.Size7:
                    he.Width = A.LineEndWidthValues.Large;
                    he.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size8:
                    he.Width = A.LineEndWidthValues.Large;
                    he.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size9:
                    he.Width = A.LineEndWidthValues.Large;
                    he.Length = A.LineEndLengthValues.Large;
                    break;
            }

            return he;
        }

        private A.TailEnd GetTailEnd()
        {
            A.TailEnd te = new A.TailEnd() { Type = this.TailEndType.Value };
            switch (this.TailEndSize)
            {
                case SLLineSizeValues.Size1:
                    te.Width = A.LineEndWidthValues.Small;
                    te.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size2:
                    te.Width = A.LineEndWidthValues.Small;
                    te.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size3:
                    te.Width = A.LineEndWidthValues.Small;
                    te.Length = A.LineEndLengthValues.Large;
                    break;
                case SLLineSizeValues.Size4:
                    te.Width = A.LineEndWidthValues.Medium;
                    te.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size5:
                    te.Width = A.LineEndWidthValues.Medium;
                    te.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size6:
                    te.Width = A.LineEndWidthValues.Medium;
                    te.Length = A.LineEndLengthValues.Large;
                    break;
                case SLLineSizeValues.Size7:
                    te.Width = A.LineEndWidthValues.Large;
                    te.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size8:
                    te.Width = A.LineEndWidthValues.Large;
                    te.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size9:
                    te.Width = A.LineEndWidthValues.Large;
                    te.Length = A.LineEndLengthValues.Large;
                    break;
            }

            return te;
        }

        internal SLLinePropertiesType Clone()
        {
            SLLinePropertiesType lpt = new SLLinePropertiesType(this.listThemeColors, this.ThrowExceptionsIfAny);
            lpt.bUseNoLine = this.bUseNoLine;
            lpt.bUseSolidLine = this.bUseSolidLine;
            lpt.SolidColor = this.SolidColor.Clone();
            lpt.bUseGradientLine = this.bUseGradientLine;
            lpt.GradientColor = this.GradientColor.Clone();
            lpt.vDashType = this.vDashType;
            lpt.HasDashType = this.HasDashType;
            lpt.vJoinType = this.vJoinType;
            lpt.HasJoinType = this.HasJoinType;
            lpt.HeadEndType = this.HeadEndType;
            lpt.HeadEndSize = this.HeadEndSize;
            lpt.TailEndType = this.TailEndType;
            lpt.TailEndSize = this.TailEndSize;
            lpt.decWidth = this.decWidth;
            lpt.HasWidth = this.HasWidth;
            lpt.vCapType = this.vCapType;
            lpt.HasCapType = this.HasCapType;
            lpt.vCompoundLineType = this.vCompoundLineType;
            lpt.HasCompoundLineType = this.HasCompoundLineType;
            lpt.Alignment = this.Alignment;

            return lpt;
        }
    }
}
