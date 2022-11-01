using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Drawing
{
    internal class SLColorTransform
    {
        internal List<System.Drawing.Color> listThemeColors;

        internal bool IsRgbColorModelHex;

        private System.Drawing.Color clrDisplayColor;
        /// <summary>
        /// This is read-only
        /// </summary>
        internal System.Drawing.Color DisplayColor
        {
            get { return clrDisplayColor; }
        }

        private System.Drawing.Color RgbColor { get; set; }
        private A.SchemeColorValues SchemeColor { get; set; }
        private decimal decTint;
        private decimal Tint
        {
            get { return decTint; }
            set
            {
                decTint = value;
                if (decTint < -1.0m) decTint = -1.0m;
                if (decTint > 1.0m) decTint = 1.0m;
            }
        }

        private decimal decTransparency;
        internal decimal Transparency
        {
            get { return decTransparency; }
            set
            {
                decTransparency = value;
                if (decTransparency < 0m) decTransparency = 0m;
                if (decTransparency > 100m) decTransparency = 100m;
            }
        }

        internal SLColorTransform(List<System.Drawing.Color> ThemeColors)
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
            this.IsRgbColorModelHex = true;
            this.clrDisplayColor = new System.Drawing.Color();
            this.RgbColor = new System.Drawing.Color();
            this.SchemeColor = A.SchemeColorValues.Light1;
            this.Tint = 0;
            this.Transparency = 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Color"></param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        internal void SetColor(System.Drawing.Color Color, decimal Transparency)
        {
            this.IsRgbColorModelHex = true;
            this.RgbColor = Color;
            this.Transparency = Transparency;

            this.clrDisplayColor = Color;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Color">The theme color used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency"></param>
        internal void SetColor(SLThemeColorIndexValues Color, double Tint, decimal Transparency)
        {
            this.IsRgbColorModelHex = false;
            switch (Color)
            {
                case SLThemeColorIndexValues.Dark1Color:
                    this.SchemeColor = A.SchemeColorValues.Dark1;
                    break;
                case SLThemeColorIndexValues.Light1Color:
                    this.SchemeColor = A.SchemeColorValues.Light1;
                    break;
                case SLThemeColorIndexValues.Dark2Color:
                    this.SchemeColor = A.SchemeColorValues.Dark2;
                    break;
                case SLThemeColorIndexValues.Light2Color:
                    this.SchemeColor = A.SchemeColorValues.Light2;
                    break;
                case SLThemeColorIndexValues.Accent1Color:
                    this.SchemeColor = A.SchemeColorValues.Accent1;
                    break;
                case SLThemeColorIndexValues.Accent2Color:
                    this.SchemeColor = A.SchemeColorValues.Accent2;
                    break;
                case SLThemeColorIndexValues.Accent3Color:
                    this.SchemeColor = A.SchemeColorValues.Accent3;
                    break;
                case SLThemeColorIndexValues.Accent4Color:
                    this.SchemeColor = A.SchemeColorValues.Accent4;
                    break;
                case SLThemeColorIndexValues.Accent5Color:
                    this.SchemeColor = A.SchemeColorValues.Accent5;
                    break;
                case SLThemeColorIndexValues.Accent6Color:
                    this.SchemeColor = A.SchemeColorValues.Accent6;
                    break;
                case SLThemeColorIndexValues.Hyperlink:
                    this.SchemeColor = A.SchemeColorValues.Hyperlink;
                    break;
                case SLThemeColorIndexValues.FollowedHyperlinkColor:
                    this.SchemeColor = A.SchemeColorValues.FollowedHyperlink;
                    break;
            }
            this.Tint = (decimal)Tint;
            this.Transparency = Transparency;

            int index = (int)Color;
            if (index >= 0 && index < this.listThemeColors.Count)
            {
                this.clrDisplayColor = System.Drawing.Color.FromArgb(255, this.listThemeColors[index]);
                if (this.Tint != 0)
                {
                    this.clrDisplayColor = SLTool.ToColor(this.clrDisplayColor, Tint);
                }
            }
        }

        internal void SetColor(A.SchemeColorValues Color, decimal Tint, decimal Transparency)
        {
            this.IsRgbColorModelHex = false;

            this.SchemeColor = Color;

            int iThemeColor = (int)SLThemeColorIndexValues.Dark1Color;
            switch (Color)
            {
                // I don't really know what to assign for Text1, Text2, Background1, Background2
                // PhClr (placeholder colour)
                case A.SchemeColorValues.Dark1:
                case A.SchemeColorValues.Text1:
                    iThemeColor = (int)SLThemeColorIndexValues.Dark1Color;
                    break;
                case A.SchemeColorValues.Light1:
                case A.SchemeColorValues.Background1:
                    iThemeColor = (int)SLThemeColorIndexValues.Light1Color;
                    break;
                case A.SchemeColorValues.Dark2:
                case A.SchemeColorValues.Text2:
                    iThemeColor = (int)SLThemeColorIndexValues.Dark2Color;
                    break;
                case A.SchemeColorValues.Light2:
                case A.SchemeColorValues.Background2:
                    iThemeColor = (int)SLThemeColorIndexValues.Light2Color;
                    break;
                case A.SchemeColorValues.PhColor:
                    iThemeColor = (int)SLThemeColorIndexValues.Accent1Color;
                    break;
                case A.SchemeColorValues.Accent1:
                    iThemeColor = (int)SLThemeColorIndexValues.Accent1Color;
                    break;
                case A.SchemeColorValues.Accent2:
                    iThemeColor = (int)SLThemeColorIndexValues.Accent2Color;
                    break;
                case A.SchemeColorValues.Accent3:
                    iThemeColor = (int)SLThemeColorIndexValues.Accent3Color;
                    break;
                case A.SchemeColorValues.Accent4:
                    iThemeColor = (int)SLThemeColorIndexValues.Accent4Color;
                    break;
                case A.SchemeColorValues.Accent5:
                    iThemeColor = (int)SLThemeColorIndexValues.Accent5Color;
                    break;
                case A.SchemeColorValues.Accent6:
                    iThemeColor = (int)SLThemeColorIndexValues.Accent6Color;
                    break;
                case A.SchemeColorValues.Hyperlink:
                    iThemeColor = (int)SLThemeColorIndexValues.Hyperlink;
                    break;
                case A.SchemeColorValues.FollowedHyperlink:
                    iThemeColor = (int)SLThemeColorIndexValues.FollowedHyperlinkColor;
                    break;
            }
            this.Tint = Tint;
            this.Transparency = Transparency;

            int index = iThemeColor;
            if (index >= 0 && index < this.listThemeColors.Count)
            {
                this.clrDisplayColor = System.Drawing.Color.FromArgb(255, this.listThemeColors[index]);
                if (this.Tint != 0)
                {
                    this.clrDisplayColor = SLTool.ToColor(this.clrDisplayColor, (double)Tint);
                }
            }
        }

        internal A.RgbColorModelHex ToRgbColorModelHex()
        {
            A.RgbColorModelHex rgb = new A.RgbColorModelHex();
            rgb.Val = string.Format("{0}{1}{2}", this.RgbColor.R.ToString("X2"), this.RgbColor.G.ToString("X2"), this.RgbColor.B.ToString("X2"));

            decimal decTint = this.Tint;

            // we don't have to do anything extra if the tint's zero.
            if (decTint < 0.0m)
            {
                decTint += 1.0m;
                decTint *= 100000m;
                rgb.Append(new A.LuminanceModulation() { Val = Convert.ToInt32(decTint) });
            }
            else if (decTint > 0.0m)
            {
                decTint *= 100000m;
                decTint = decimal.Floor(decTint);
                rgb.Append(new A.LuminanceModulation() { Val = Convert.ToInt32(100000m - decTint) });
                rgb.Append(new A.LuminanceOffset() { Val = Convert.ToInt32(decTint) });
            }

            int iAlpha = SLA.SLDrawingTool.CalculateAlpha(Transparency);
            // if >= 100000, then transparency was 0 (or negative),
            // then we don't have to append the Alpha class
            if (iAlpha < 100000)
            {
                rgb.Append(new A.Alpha() { Val = iAlpha });
            }

            return rgb;
        }

        internal A.SchemeColor ToSchemeColor()
        {
            A.SchemeColor sclr = new A.SchemeColor();
            sclr.Val = this.SchemeColor;

            decimal decTint = this.Tint;

            // we don't have to do anything extra if the tint's zero.
            if (decTint < 0.0m)
            {
                decTint += 1.0m;
                decTint *= 100000m;
                sclr.Append(new A.LuminanceModulation() { Val = Convert.ToInt32(decTint) });
            }
            else if (decTint > 0.0m)
            {
                decTint *= 100000m;
                decTint = decimal.Floor(decTint);
                sclr.Append(new A.LuminanceModulation() { Val = Convert.ToInt32(100000m - decTint) });
                sclr.Append(new A.LuminanceOffset() { Val = Convert.ToInt32(decTint) });
            }

            int iAlpha = SLA.SLDrawingTool.CalculateAlpha(Transparency);
            // if >= 100000, then transparency was 0 (or negative),
            // then we don't have to append the Alpha class
            if (iAlpha < 100000)
            {
                sclr.Append(new A.Alpha() { Val = iAlpha });
            }

            return sclr;
        }

        internal SLColorTransform Clone()
        {
            SLColorTransform clr = new SLColorTransform(this.listThemeColors);
            clr.IsRgbColorModelHex = this.IsRgbColorModelHex;
            clr.clrDisplayColor = this.clrDisplayColor;
            clr.RgbColor = this.RgbColor;
            clr.SchemeColor = this.SchemeColor;
            clr.decTint = this.decTint;
            clr.decTransparency = this.decTransparency;

            return clr;
        }
    }
}
