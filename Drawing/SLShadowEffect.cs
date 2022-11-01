using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Drawing
{
    /// <summary>
    /// Encapsulates properties and methods for specifying shadow effects.
    /// This simulates the DocumentFormat.OpenXml.Drawing.InnerShadow and DocumentFormat.OpenXml.Drawing.OuterShadow classes.
    /// </summary>
    public class SLShadowEffect
    {
        internal List<System.Drawing.Color> listThemeColors;

        /// <summary>
        /// doubles as HasShadow variable
        /// </summary>
        internal bool? IsInnerShadow { get; set; }

        internal SLA.SLColorTransform InnerShadowColor { get; set; }
        internal decimal InnerShadowBlurRadius { get; set; }
        internal decimal InnerShadowDistance { get; set; }
        internal decimal InnerShadowDirection { get; set; }

        internal SLA.SLColorTransform OuterShadowColor { get; set; }
        internal decimal OuterShadowBlurRadius { get; set; }
        internal decimal OuterShadowDistance { get; set; }
        internal decimal OuterShadowDirection { get; set; }
        internal decimal OuterShadowHorizontalRatio { get; set; }
        internal decimal OuterShadowVerticalRatio { get; set; }
        internal decimal OuterShadowHorizontalSkew { get; set; }
        internal decimal OuterShadowVerticalSkew { get; set; }
        internal A.RectangleAlignmentValues OuterShadowAlignment { get; set; }
        internal bool OuterShadowRotateWithShape { get; set; }

        /// <summary>
        /// The shadow color.
        /// </summary>
        public System.Drawing.Color Color
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                    {
                        return this.InnerShadowColor.DisplayColor;
                    }
                    else
                    {
                        return this.OuterShadowColor.DisplayColor;
                    }
                }
                else
                {
                    return new System.Drawing.Color();
                }
            }
        }

        /// <summary>
        /// Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.
        /// </summary>
        public decimal Transparency
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                    {
                        return this.InnerShadowColor.Transparency;
                    }
                    else
                    {
                        return this.OuterShadowColor.Transparency;
                    }
                }
                else
                {
                    return 0;
                }
            }
            set
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                    {
                        this.InnerShadowColor.Transparency = value;
                    }
                    else
                    {
                        this.OuterShadowColor.Transparency = value;
                    }
                }
            }
        }

        /// <summary>
        /// Specifies the size of the shadow in percentage. While there's no restriction in range, consider a range of 1% to 200%. Accurate to 1/1000th of a percent.
        /// </summary>
        public decimal Size
        {
            get
            {
                return this.OuterShadowHorizontalRatio;
            }
            set
            {
                decimal dec = value;
                this.OuterShadowHorizontalRatio = dec;
                this.OuterShadowVerticalRatio = dec;
            }
        }

        /// <summary>
        /// Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point.
        /// </summary>
        public decimal Blur
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                    {
                        return this.InnerShadowBlurRadius;
                    }
                    else
                    {
                        return this.OuterShadowBlurRadius;
                    }
                }
                else
                {
                    return 0;
                }
            }
            set
            {
                if (IsInnerShadow != null)
                {
                    decimal dec = value;
                    if (dec < 0m) dec = 0m;
                    if (dec > 100m) dec = 100m;

                    if (IsInnerShadow.Value)
                    {
                        this.InnerShadowBlurRadius = dec;
                    }
                    else
                    {
                        this.OuterShadowBlurRadius = dec;
                    }
                }
            }
        }

        /// <summary>
        /// Angle of shadow projection, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.
        /// </summary>
        public decimal Angle
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                    {
                        return this.InnerShadowDirection;
                    }
                    else
                    {
                        return this.OuterShadowDirection;
                    }
                }
                else
                {
                    return 0;
                }
            }
            set
            {
                if (IsInnerShadow != null)
                {
                    decimal dec = value;
                    if (dec < 0m) dec = 0m;
                    if (dec >= 360m) dec = 359.9m;

                    if (IsInnerShadow.Value)
                    {
                        this.InnerShadowDirection = dec;
                    }
                    else
                    {
                        this.InnerShadowDirection = dec;
                    }
                }
            }
        }

        /// <summary>
        /// Distance of shadow away from source object, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.
        /// </summary>
        public decimal Distance
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                    {
                        return this.InnerShadowDistance;
                    }
                    else
                    {
                        return this.OuterShadowDistance;
                    }
                }
                else
                {
                    return 0;
                }
            }
            set
            {
                if (IsInnerShadow != null)
                {
                    decimal dec = value;
                    if (dec < 0m) dec = 0m;
                    if (dec > 200m) dec = 200m;

                    if (IsInnerShadow.Value)
                    {
                        this.InnerShadowDistance = dec;
                    }
                    else
                    {
                        this.OuterShadowDistance = dec;
                    }
                }
            }
        }

        internal SLShadowEffect(List<System.Drawing.Color> ThemeColors)
        {
            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }

            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.IsInnerShadow = null;

            this.InnerShadowColor = new SLColorTransform(this.listThemeColors);
            this.InnerShadowBlurRadius = 0;
            this.InnerShadowDistance = 0;
            this.InnerShadowDirection = 0;

            this.OuterShadowColor = new SLColorTransform(this.listThemeColors);
            this.OuterShadowBlurRadius = 0;
            this.OuterShadowDistance = 0;
            this.OuterShadowDirection = 0;
            this.OuterShadowHorizontalRatio = 100;
            this.OuterShadowVerticalRatio = 100;
            this.OuterShadowHorizontalSkew = 0;
            this.OuterShadowVerticalSkew = 0;
            this.OuterShadowAlignment = A.RectangleAlignmentValues.Bottom;
            this.OuterShadowRotateWithShape = true;
        }

        /// <summary>
        /// Set the shadow color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetShadowColor(System.Drawing.Color Color, decimal Transparency)
        {
            if (IsInnerShadow != null)
            {
                if (IsInnerShadow.Value)
                {
                    this.InnerShadowColor.SetColor(Color, Transparency);
                }
                else
                {
                    this.OuterShadowColor.SetColor(Color, Transparency);
                }
            }
        }

        /// <summary>
        /// Set the shadow color.
        /// </summary>
        /// <param name="Color">The theme color used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetShadowColor(SLThemeColorIndexValues Color, double Tint, decimal Transparency)
        {
            if (IsInnerShadow != null)
            {
                if (IsInnerShadow.Value)
                {
                    this.InnerShadowColor.SetColor(Color, Tint, Transparency);
                }
                else
                {
                    this.OuterShadowColor.SetColor(Color, Tint, Transparency);
                }
            }
        }

        /// <summary>
        /// Set a shadow using a preset.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        public void SetPreset(SLShadowPresetValues Preset)
        {
            System.Drawing.Color clr = System.Drawing.Color.FromArgb(0, 0, 0);

            switch (Preset)
            {
                case SLShadowPresetValues.None:
                    this.SetAllNull();
                    break;
                case SLShadowPresetValues.OuterDiagonalBottomRight:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 60);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 100;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 4;
                    this.OuterShadowDirection = 45;
                    this.OuterShadowDistance = 3;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.TopLeft;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterBottom:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 60);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 100;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 4;
                    this.OuterShadowDirection = 90;
                    this.OuterShadowDistance = 3;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.Top;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterDiagonalBottomLeft:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 60);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 100;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 4;
                    this.OuterShadowDirection = 135;
                    this.OuterShadowDistance = 3;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.TopRight;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterRight:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 60);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 100;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 4;
                    this.OuterShadowDirection = 0;
                    this.OuterShadowDistance = 3;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.Left;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterCenter:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 60);
                    this.OuterShadowHorizontalRatio = 102;
                    this.OuterShadowVerticalRatio = 102;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 5;
                    this.OuterShadowDirection = 0;
                    this.OuterShadowDistance = 0;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.Center;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterLeft:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 60);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 100;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 4;
                    this.OuterShadowDirection = 180;
                    this.OuterShadowDistance = 3;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.Right;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterDiagonalTopRight:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 60);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 100;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 4;
                    this.OuterShadowDirection = 315;
                    this.OuterShadowDistance = 3;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.BottomLeft;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterTop:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 60);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 100;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 4;
                    this.OuterShadowDirection = 270;
                    this.OuterShadowDistance = 3;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.Bottom;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterDiagonalTopLeft:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 60);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 100;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 4;
                    this.OuterShadowDirection = 225;
                    this.OuterShadowDistance = 3;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.BottomRight;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.InnerDiagonalTopLeft:
                    this.IsInnerShadow = true;
                    this.InnerShadowColor.SetColor(clr, 50);
                    this.InnerShadowBlurRadius = 5;
                    this.InnerShadowDirection = 225;
                    this.InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerTop:
                    this.IsInnerShadow = true;
                    this.InnerShadowColor.SetColor(clr, 50);
                    this.InnerShadowBlurRadius = 5;
                    this.InnerShadowDirection = 270;
                    this.InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerDiagonalTopRight:
                    this.IsInnerShadow = true;
                    this.InnerShadowColor.SetColor(clr, 50);
                    this.InnerShadowBlurRadius = 5;
                    this.InnerShadowDirection = 315;
                    this.InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerLeft:
                    this.IsInnerShadow = true;
                    this.InnerShadowColor.SetColor(clr, 50);
                    this.InnerShadowBlurRadius = 5;
                    this.InnerShadowDirection = 180;
                    this.InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerCenter:
                    this.IsInnerShadow = true;
                    this.InnerShadowColor.SetColor(clr, 0);
                    this.InnerShadowBlurRadius = 9;
                    this.InnerShadowDirection = 0;
                    this.InnerShadowDistance = 0;
                    break;
                case SLShadowPresetValues.InnerRight:
                    this.IsInnerShadow = true;
                    this.InnerShadowColor.SetColor(clr, 50);
                    this.InnerShadowBlurRadius = 5;
                    this.InnerShadowDirection = 0;
                    this.InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerDiagonalBottomLeft:
                    this.IsInnerShadow = true;
                    this.InnerShadowColor.SetColor(clr, 50);
                    this.InnerShadowBlurRadius = 5;
                    this.InnerShadowDirection = 135;
                    this.InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerBottom:
                    this.IsInnerShadow = true;
                    this.InnerShadowColor.SetColor(clr, 50);
                    this.InnerShadowBlurRadius = 5;
                    this.InnerShadowDirection = 90;
                    this.InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerDiagonalBottomRight:
                    this.IsInnerShadow = true;
                    this.InnerShadowColor.SetColor(clr, 50);
                    this.InnerShadowBlurRadius = 5;
                    this.InnerShadowDirection = 45;
                    this.InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.PerspectiveDiagonalUpperLeft:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 80);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 23;
                    this.OuterShadowHorizontalSkew = 20;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 6;
                    this.OuterShadowDirection = 225;
                    this.OuterShadowDistance = 0;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.BottomRight;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.PerspectiveDiagonalUpperRight:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 80);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = 23;
                    this.OuterShadowHorizontalSkew = -20;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 6;
                    this.OuterShadowDirection = 315;
                    this.OuterShadowDistance = 0;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.BottomLeft;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.PerspectiveBelow:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 85);
                    this.OuterShadowHorizontalRatio = 90;
                    this.OuterShadowVerticalRatio = 100;
                    this.OuterShadowHorizontalSkew = 0;
                    this.OuterShadowVerticalSkew = -0.3166667m;
                    this.OuterShadowBlurRadius = 12;
                    this.OuterShadowDirection = 90;
                    this.OuterShadowDistance = 25;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.Bottom;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.PerspectiveDiagonalLowerLeft:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 80);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = -23;
                    this.OuterShadowHorizontalSkew = 13.34m;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 6;
                    this.OuterShadowDirection = 135;
                    this.OuterShadowDistance = 1;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.BottomRight;
                    this.OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.PerspectiveDiagonalLowerRight:
                    this.IsInnerShadow = false;
                    this.OuterShadowColor.SetColor(clr, 80);
                    this.OuterShadowHorizontalRatio = 100;
                    this.OuterShadowVerticalRatio = -23;
                    this.OuterShadowHorizontalSkew = -13.34m;
                    this.OuterShadowVerticalSkew = 0;
                    this.OuterShadowBlurRadius = 6;
                    this.OuterShadowDirection = 45;
                    this.OuterShadowDistance = 1;
                    this.OuterShadowAlignment = A.RectangleAlignmentValues.BottomLeft;
                    this.OuterShadowRotateWithShape = false;
                    break;
            }
        }

        // TODO overload setting of inner and outer shadow functions here

        internal A.InnerShadow ToInnerShadow()
        {
            A.InnerShadow ishad = new A.InnerShadow();
            if (this.InnerShadowColor.IsRgbColorModelHex)
            {
                ishad.RgbColorModelHex = this.InnerShadowColor.ToRgbColorModelHex();
            }
            else
            {
                ishad.SchemeColor = this.InnerShadowColor.ToSchemeColor();
            }

            if (this.InnerShadowBlurRadius != 0)
            {
                ishad.BlurRadius = SLA.SLDrawingTool.CalculatePositiveCoordinate(this.InnerShadowBlurRadius);
            }

            if (this.InnerShadowDistance != 0)
            {
                ishad.Distance = SLA.SLDrawingTool.CalculatePositiveCoordinate(this.InnerShadowDistance);
            }

            if (this.InnerShadowDirection != 0)
            {
                ishad.Direction = SLA.SLDrawingTool.CalculatePositiveFixedAngle(this.InnerShadowDirection);
            }

            return ishad;
        }

        internal A.OuterShadow ToOuterShadow()
        {
            A.OuterShadow os = new A.OuterShadow();

            if (this.OuterShadowColor.IsRgbColorModelHex)
            {
                os.RgbColorModelHex = this.OuterShadowColor.ToRgbColorModelHex();
            }
            else
            {
                os.SchemeColor = this.OuterShadowColor.ToSchemeColor();
            }

            if (this.OuterShadowBlurRadius != 0)
            {
                os.BlurRadius = SLA.SLDrawingTool.CalculatePositiveCoordinate(this.OuterShadowBlurRadius);
            }

            if (this.OuterShadowDistance != 0)
            {
                os.Distance = SLA.SLDrawingTool.CalculatePositiveCoordinate(this.OuterShadowDistance);
            }

            if (this.OuterShadowDirection != 0)
            {
                os.Direction = SLA.SLDrawingTool.CalculatePositiveFixedAngle(this.OuterShadowDirection);
            }

            if (this.OuterShadowHorizontalRatio != 100m)
            {
                os.HorizontalRatio = SLA.SLDrawingTool.CalculatePercentage(this.OuterShadowHorizontalRatio);
            }

            if (this.OuterShadowVerticalRatio != 100m)
            {
                os.VerticalRatio = SLA.SLDrawingTool.CalculatePercentage(this.OuterShadowVerticalRatio);
            }

            if (this.OuterShadowHorizontalSkew != 0m)
            {
                os.HorizontalSkew = SLA.SLDrawingTool.CalculateFixedAngle(this.OuterShadowHorizontalSkew);
            }

            if (this.OuterShadowVerticalSkew != 0m)
            {
                os.VerticalSkew = SLA.SLDrawingTool.CalculateFixedAngle(this.OuterShadowVerticalSkew);
            }

            if (this.OuterShadowAlignment != A.RectangleAlignmentValues.Bottom) os.Alignment = this.OuterShadowAlignment;

            if (!this.OuterShadowRotateWithShape) os.RotateWithShape = this.OuterShadowRotateWithShape;

            return os;
        }

        internal SLShadowEffect Clone()
        {
            SLShadowEffect se = new SLShadowEffect(this.listThemeColors);
            se.IsInnerShadow = this.IsInnerShadow;
            se.InnerShadowColor = this.InnerShadowColor.Clone();
            se.InnerShadowBlurRadius = this.InnerShadowBlurRadius;
            se.InnerShadowDistance = this.InnerShadowDistance;
            se.InnerShadowDirection = this.InnerShadowDirection;
            se.OuterShadowColor = this.OuterShadowColor.Clone();
            se.OuterShadowBlurRadius = this.OuterShadowBlurRadius;
            se.OuterShadowDistance = this.OuterShadowDistance;
            se.OuterShadowDirection = this.OuterShadowDirection;
            se.OuterShadowHorizontalRatio = this.OuterShadowHorizontalRatio;
            se.OuterShadowVerticalRatio = this.OuterShadowVerticalRatio;
            se.OuterShadowHorizontalSkew = this.OuterShadowHorizontalSkew;
            se.OuterShadowVerticalSkew = this.OuterShadowVerticalSkew;
            se.OuterShadowAlignment = this.OuterShadowAlignment;
            se.OuterShadowRotateWithShape = this.OuterShadowRotateWithShape;

            return se;
        }
    }
}
