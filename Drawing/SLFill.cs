using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Drawing
{
    internal enum SLFillType
    {
        Automatic = 0,
        NoFill,
        SolidFill,
        GradientFill,
        BlipFill,
        PatternFill
    }

    /// <summary>
    /// Encapsulates properties and methods for specifying fill effects.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Fill class.
    /// </summary>
    public class SLFill
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal bool ThrowExceptionsIfAny { get; set; }

        internal SLFillType Type;
        internal bool HasFill
        {
            get { return Type != SLFillType.Automatic ? true : false; }
        }

        internal SLColorTransform SolidColor { get; set; }

        internal SLGradientFill GradientColor { get; set; }

        internal string BlipFileName { get; set; }
        internal string BlipRelationshipID { get; set; }
        internal bool BlipTile { get; set; }
        private decimal decBlipLeftOffset;
        internal decimal BlipLeftOffset
        {
            get { return decBlipLeftOffset; }
            set
            {
                decBlipLeftOffset = value;
                if (decBlipLeftOffset < -100m) decBlipLeftOffset = -100m;
                if (decBlipLeftOffset > 100m) decBlipLeftOffset = 100m;
            }
        }
        private decimal decBlipRightOffset;
        internal decimal BlipRightOffset
        {
            get { return decBlipRightOffset; }
            set
            {
                decBlipRightOffset = value;
                if (decBlipRightOffset < -100m) decBlipRightOffset = -100m;
                if (decBlipRightOffset > 100m) decBlipRightOffset = 100m;
            }
        }
        private decimal decBlipTopOffset;
        internal decimal BlipTopOffset
        {
            get { return decBlipTopOffset; }
            set
            {
                decBlipTopOffset = value;
                if (decBlipTopOffset < -100m) decBlipTopOffset = -100m;
                if (decBlipTopOffset > 100m) decBlipTopOffset = 100m;
            }
        }
        private decimal decBlipBottomOffset;
        internal decimal BlipBottomOffset
        {
            get { return decBlipBottomOffset; }
            set
            {
                decBlipBottomOffset = value;
                if (decBlipBottomOffset < -100m) decBlipBottomOffset = -100m;
                if (decBlipBottomOffset > 100m) decBlipBottomOffset = 100m;
            }
        }
        private decimal decBlipOffsetX;
        internal decimal BlipOffsetX
        {
            get { return decBlipOffsetX; }
            set
            {
                decBlipOffsetX = value;
                if (decBlipOffsetX < -1584m) decBlipOffsetX = -1584m;
                if (decBlipOffsetX > 1584m) decBlipOffsetX = 1584m;
            }
        }
        private decimal decBlipOffsetY;
        internal decimal BlipOffsetY
        {
            get { return decBlipOffsetY; }
            set
            {
                decBlipOffsetY = value;
                if (decBlipOffsetY < -1584m) decBlipOffsetY = -1584m;
                if (decBlipOffsetY > 1584m) decBlipOffsetY = 1584m;
            }
        }
        private decimal decBlipScaleX;
        internal decimal BlipScaleX
        {
            get { return decBlipScaleX; }
            set
            {
                decBlipScaleX = value;
                if (decBlipScaleX < 0m) decBlipScaleX = 0m;
                if (decBlipScaleX > 100m) decBlipScaleX = 100m;
            }
        }
        private decimal decBlipScaleY;
        internal decimal BlipScaleY
        {
            get { return decBlipScaleY; }
            set
            {
                decBlipScaleY = value;
                if (decBlipScaleY < 0m) decBlipScaleY = 0m;
                if (decBlipScaleY > 100m) decBlipScaleY = 100m;
            }
        }
        internal A.RectangleAlignmentValues BlipAlignment { get; set; }
        internal A.TileFlipValues BlipMirrorType { get; set; }
        private decimal decBlipTransparency;
        internal decimal BlipTransparency
        {
            get { return decBlipTransparency; }
            set
            {
                decBlipTransparency = value;
                if (decBlipTransparency < 0m) decBlipTransparency = 0m;
                if (decBlipTransparency > 100m) decBlipTransparency = 100m;
            }
        }
        internal uint? BlipDpi { get; set; }
        internal bool? BlipRotateWithShape { get; set; }

        internal A.PresetPatternValues PatternPreset { get; set; }
        internal SLColorTransform PatternForegroundColor { get; set; }
        internal SLColorTransform PatternBackgroundColor { get; set; }

        internal SLFill(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
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
            this.Type = SLFillType.Automatic;
            this.SolidColor = new SLColorTransform(this.listThemeColors);
            this.GradientColor = new SLGradientFill(this.listThemeColors, this.ThrowExceptionsIfAny);
            this.BlipFileName = string.Empty;
            this.BlipRelationshipID = string.Empty;
            this.BlipTile = true;
            this.BlipLeftOffset = 0;
            this.BlipRightOffset = 0;
            this.BlipTopOffset = 0;
            this.BlipBottomOffset = 0;
            this.BlipOffsetX = 0;
            this.BlipOffsetY = 0;
            this.BlipScaleX = 100;
            this.BlipScaleY = 100;
            this.BlipAlignment = A.RectangleAlignmentValues.TopLeft;
            this.BlipMirrorType = A.TileFlipValues.None;
            this.BlipTransparency = 0;
            this.BlipDpi = null;
            this.BlipRotateWithShape = null;
            this.PatternForegroundColor = new SLColorTransform(this.listThemeColors);
            this.PatternBackgroundColor = new SLColorTransform(this.listThemeColors);
        }

        /// <summary>
        /// Set the fill to automatic.
        /// </summary>
        public void SetAutomaticFill()
        {
            this.Type = SLFillType.Automatic;
        }

        /// <summary>
        /// Set no fill.
        /// </summary>
        public void SetNoFill()
        {
            this.Type = SLFillType.NoFill;
        }

        /// <summary>
        /// Set a solid fill.
        /// </summary>
        /// <param name="FillColor">The color used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetSolidFill(System.Drawing.Color FillColor, decimal Transparency)
        {
            this.Type = SLFillType.SolidFill;
            this.SolidColor.SetColor(FillColor, Transparency);
        }

        /// <summary>
        /// Set a solid fill.
        /// </summary>
        /// <param name="FillColor">The theme color used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetSolidFill(SLThemeColorIndexValues FillColor, double Tint, decimal Transparency)
        {
            this.Type = SLFillType.SolidFill;
            this.SolidColor.SetColor(FillColor, Tint, Transparency);
        }

        internal void SetSolidFill(A.SchemeColorValues FillColor, decimal Tint, decimal Transparency)
        {
            this.Type = SLFillType.SolidFill;
            this.SolidColor.SetColor(FillColor, Tint, Transparency);
        }

        /// <summary>
        /// Set a linear gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Angle">The interpolation angle ranging from 0 degrees to 359.9 degrees. 0 degrees mean from left to right, 90 degrees mean from top to bottom, 180 degrees mean from right to left and 270 degrees mean from bottom to top. Accurate to 1/60000 of a degree.</param>
        public void SetLinearGradient(SLGradientPresetValues Preset, decimal Angle)
        {
            this.Type = SLFillType.GradientFill;
            this.GradientColor.SetLinearGradient(Preset, Angle);
        }

        /// <summary>
        /// Set a radial gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Direction">The radial gradient direction.</param>
        public void SetRadialGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            this.Type = SLFillType.GradientFill;
            this.GradientColor.SetRadialGradient(Preset, Direction);
        }

        /// <summary>
        /// Set a rectangular gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Direction">The rectangular gradient direction.</param>
        public void SetRectangularGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            this.Type = SLFillType.GradientFill;
            this.GradientColor.SetRectangularGradient(Preset, Direction);
        }

        /// <summary>
        /// Set a path gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        public void SetPathGradient(SLGradientPresetValues Preset)
        {
            this.Type = SLFillType.GradientFill;
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
        /// Set a picture fill. This stretches the picture.
        /// </summary>
        /// <param name="PictureFileName">The file name of the image/picture used.</param>
        /// <param name="LeftOffset">The left offset in percentage. A suggested range is -100% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="RightOffset">The right offset in percentage. A suggested range is -100% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="TopOffset">The top offset in percentage. A suggested range is -100% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="BottomOffset">The bottom offset in percentage. A suggested range is -100% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Transparency">Transparency of the picture ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetPictureFill(string PictureFileName, decimal LeftOffset, decimal RightOffset, decimal TopOffset, decimal BottomOffset, decimal Transparency)
        {
            this.Type = SLFillType.BlipFill;
            this.BlipTile = false;
            this.BlipFileName = PictureFileName;
            this.BlipLeftOffset = LeftOffset;
            this.BlipRightOffset = RightOffset;
            this.BlipTopOffset = TopOffset;
            this.BlipBottomOffset = BottomOffset;
            this.BlipTransparency = Transparency;
        }

        /// <summary>
        /// Set a picture fill. This tiles the picture.
        /// </summary>
        /// <param name="PictureFileName">The file name of the image/picture used.</param>
        /// <param name="OffsetX">Horizontal offset ranging from -2147483648 pt to 2147483647 pt. However a suggested range is -1585pt to 1584pt. Accurate to 1/12700 of a point.</param>
        /// <param name="OffsetY">Vertical offset ranging from -2147483648 pt to 2147483647 pt. However a suggested range is -1585pt to 1584pt. Accurate to 1/12700 of a point.</param>
        /// <param name="ScaleX">Horizontal scale in percentage. A suggested range is 0% to 100%.</param>
        /// <param name="ScaleY">Vertical scale in percentage. A suggested range is 0% to 100%.</param>
        /// <param name="Alignment">Picture alignment.</param>
        /// <param name="MirrorType">Picture mirror type.</param>
        /// <param name="Transparency">Transparency of the picture ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetPictureFill(string PictureFileName, decimal OffsetX, decimal OffsetY, decimal ScaleX, decimal ScaleY, A.RectangleAlignmentValues Alignment, A.TileFlipValues MirrorType, decimal Transparency)
        {
            this.Type = SLFillType.BlipFill;
            this.BlipTile = true;
            this.BlipFileName = PictureFileName;
            this.BlipOffsetX = OffsetX;
            this.BlipOffsetY = OffsetY;
            this.BlipScaleX = ScaleX;
            this.BlipScaleY = ScaleY;
            this.BlipAlignment = Alignment;
            this.BlipMirrorType = MirrorType;
            this.BlipTransparency = Transparency;
        }

        /// <summary>
        /// Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, System.Drawing.Color ForegroundColor, System.Drawing.Color BackgroundColor)
        {
            this.Type = SLFillType.PatternFill;
            this.PatternPreset = PresetPattern;
            this.PatternForegroundColor.SetColor(ForegroundColor, 0);
            this.PatternBackgroundColor.SetColor(BackgroundColor, 0);
        }

        /// <summary>
        /// Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, System.Drawing.Color ForegroundColor, SLThemeColorIndexValues BackgroundColorTheme)
        {
            this.Type = SLFillType.PatternFill;
            this.PatternPreset = PresetPattern;
            this.PatternForegroundColor.SetColor(ForegroundColor, 0);
            this.PatternBackgroundColor.SetColor(BackgroundColorTheme, 0, 0);
        }

        /// <summary>
        /// Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, System.Drawing.Color ForegroundColor, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            this.Type = SLFillType.PatternFill;
            this.PatternPreset = PresetPattern;
            this.PatternForegroundColor.SetColor(ForegroundColor, 0);
            this.PatternBackgroundColor.SetColor(BackgroundColorTheme, BackgroundColorTint, 0);
        }

        /// <summary>
        /// Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme, System.Drawing.Color BackgroundColor)
        {
            this.Type = SLFillType.PatternFill;
            this.PatternPreset = PresetPattern;
            this.PatternForegroundColor.SetColor(ForegroundColorTheme, 0, 0);
            this.PatternBackgroundColor.SetColor(BackgroundColor, 0);
        }

        /// <summary>
        /// Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme, SLThemeColorIndexValues BackgroundColorTheme)
        {
            this.Type = SLFillType.PatternFill;
            this.PatternPreset = PresetPattern;
            this.PatternForegroundColor.SetColor(ForegroundColorTheme, 0, 0);
            this.PatternBackgroundColor.SetColor(BackgroundColorTheme, 0, 0);
        }

        /// <summary>
        /// Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            this.Type = SLFillType.PatternFill;
            this.PatternPreset = PresetPattern;
            this.PatternForegroundColor.SetColor(ForegroundColorTheme, 0, 0);
            this.PatternBackgroundColor.SetColor(BackgroundColorTheme, BackgroundColorTint, 0);
        }

        /// <summary>
        /// Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme, double ForegroundColorTint, System.Drawing.Color BackgroundColor)
        {
            this.Type = SLFillType.PatternFill;
            this.PatternPreset = PresetPattern;
            this.PatternForegroundColor.SetColor(ForegroundColorTheme, ForegroundColorTint, 0);
            this.PatternBackgroundColor.SetColor(BackgroundColor, 0);
        }

        /// <summary>
        /// Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme, double ForegroundColorTint, SLThemeColorIndexValues BackgroundColorTheme)
        {
            this.Type = SLFillType.PatternFill;
            this.PatternPreset = PresetPattern;
            this.PatternForegroundColor.SetColor(ForegroundColorTheme, ForegroundColorTint, 0);
            this.PatternBackgroundColor.SetColor(BackgroundColorTheme, 0, 0);
        }

        /// <summary>
        /// Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme, double ForegroundColorTint, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            this.Type = SLFillType.PatternFill;
            this.PatternPreset = PresetPattern;
            this.PatternForegroundColor.SetColor(ForegroundColorTheme, ForegroundColorTint, 0);
            this.PatternBackgroundColor.SetColor(BackgroundColorTheme, BackgroundColorTint, 0);
        }

        internal OpenXmlElement ToFill()
        {
            OpenXmlElement oxe = new A.NoFill();

            if (this.Type == SLFillType.NoFill)
            {
                return new A.NoFill();
            }
            else if (this.Type == SLFillType.SolidFill)
            {
                A.SolidFill sf = new A.SolidFill();
                if (this.SolidColor.IsRgbColorModelHex)
                {
                    sf.RgbColorModelHex = this.SolidColor.ToRgbColorModelHex();
                }
                else
                {
                    sf.SchemeColor = this.SolidColor.ToSchemeColor();
                }
                return sf;
            }
            else if (this.Type == SLFillType.GradientFill)
            {
                return this.GradientColor.ToGradientFill();
            }
            else if (this.Type == SLFillType.BlipFill)
            {
                A.BlipFill bf = new A.BlipFill();
                if (this.BlipDpi != null) bf.Dpi = this.BlipDpi.Value;
                if (this.BlipRotateWithShape != null) bf.RotateWithShape = this.BlipRotateWithShape.Value;

                bf.Blip = new A.Blip();
                bf.Blip.Embed = this.BlipRelationshipID;
                if (this.BlipTransparency > 0m)
                {
                    bf.Blip.Append(new A.AlphaModulationFixed() { Amount = SLA.SLDrawingTool.CalculateAlpha(this.BlipTransparency) });
                }
                bf.Append(new A.SourceRectangle());
                if (this.BlipTile)
                {
                    bf.Append(new A.Tile()
                    {
                        HorizontalOffset = SLA.SLDrawingTool.CalculateCoordinate(this.BlipOffsetX),
                        VerticalOffset = SLA.SLDrawingTool.CalculateCoordinate(this.BlipOffsetY),
                        HorizontalRatio = SLA.SLDrawingTool.CalculatePercentage(this.BlipScaleX),
                        VerticalRatio = SLA.SLDrawingTool.CalculatePercentage(this.BlipScaleY),
                        Flip = this.BlipMirrorType,
                        Alignment = this.BlipAlignment
                    });
                }
                else
                {
                    bf.Append(new A.Stretch()
                    {
                        FillRectangle = new A.FillRectangle()
                        {
                            Left = SLA.SLDrawingTool.CalculatePercentage(this.BlipLeftOffset),
                            Top = SLA.SLDrawingTool.CalculatePercentage(this.BlipTopOffset),
                            Right = SLA.SLDrawingTool.CalculatePercentage(this.BlipRightOffset),
                            Bottom = SLA.SLDrawingTool.CalculatePercentage(this.BlipBottomOffset)
                        }
                    });
                }
                return bf;
            }
            else if (this.Type == SLFillType.PatternFill)
            {
                A.PatternFill pf = new A.PatternFill();
                pf.Preset = A.PresetPatternValues.Trellis;

                pf.ForegroundColor = new A.ForegroundColor();
                if (this.PatternForegroundColor.IsRgbColorModelHex)
                {
                    pf.ForegroundColor.RgbColorModelHex = this.PatternForegroundColor.ToRgbColorModelHex();
                }
                else
                {
                    pf.ForegroundColor.SchemeColor = this.PatternForegroundColor.ToSchemeColor();
                }

                pf.BackgroundColor = new A.BackgroundColor();
                if (this.PatternBackgroundColor.IsRgbColorModelHex)
                {
                    pf.BackgroundColor.RgbColorModelHex = this.PatternBackgroundColor.ToRgbColorModelHex();
                }
                else
                {
                    pf.BackgroundColor.SchemeColor = this.PatternBackgroundColor.ToSchemeColor();
                }

                return pf;
            }

            return oxe;
        }

        internal SLFill Clone()
        {
            SLFill fill = new SLFill(this.listThemeColors, this.ThrowExceptionsIfAny);
            fill.Type = this.Type;
            fill.SolidColor = this.SolidColor.Clone();
            fill.GradientColor = this.GradientColor.Clone();
            fill.BlipFileName = this.BlipFileName;
            fill.BlipRelationshipID = this.BlipRelationshipID;
            fill.BlipTile = this.BlipTile;
            fill.decBlipLeftOffset = this.decBlipLeftOffset;
            fill.decBlipRightOffset = this.decBlipRightOffset;
            fill.decBlipTopOffset = this.decBlipTopOffset;
            fill.decBlipBottomOffset = this.decBlipBottomOffset;
            fill.decBlipOffsetX = this.decBlipOffsetX;
            fill.decBlipOffsetY = this.decBlipOffsetY;
            fill.decBlipScaleX = this.decBlipScaleX;
            fill.decBlipScaleY = this.decBlipScaleY;
            fill.BlipAlignment = this.BlipAlignment;
            fill.BlipMirrorType = this.BlipMirrorType;
            fill.decBlipTransparency = this.decBlipTransparency;
            fill.BlipDpi = this.BlipDpi;
            fill.BlipRotateWithShape = this.BlipRotateWithShape;
            fill.PatternPreset = this.PatternPreset;
            fill.PatternForegroundColor = this.PatternForegroundColor.Clone();
            fill.PatternBackgroundColor = this.PatternBackgroundColor.Clone();

            return fill;
        }
    }
}
