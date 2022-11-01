using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing
{
    internal class SLGradientStop
    {
        internal SLColorTransform Color { get; set; }
        private decimal decPosition;
        /// <summary>
        /// The position in percentage ranging from 0% to 100%. Accurate to 1/1000 of a percent.
        /// </summary>
        internal decimal Position
        {
            get { return decPosition; }
            set
            {
                decPosition = value;
                if (decPosition < 0m) decPosition = 0m;
                if (decPosition > 100m) decPosition = 100m;
            }
        }

        internal SLGradientStop(List<System.Drawing.Color> ThemeColors)
        {
            this.Color = new SLColorTransform(ThemeColors);
            this.Position = 0m;
        }

        internal SLGradientStop(List<System.Drawing.Color> ThemeColors, string HexColor, decimal Position, bool ThrowExceptionsIfAny)
        {
            this.Color = new SLColorTransform(ThemeColors);
            this.Position = Position;

            System.Drawing.Color clr = new System.Drawing.Color();
            try
            {
                clr = System.Drawing.Color.FromArgb(int.Parse(HexColor, System.Globalization.NumberStyles.HexNumber));
            }
            catch (Exception e)
            {
                if (ThrowExceptionsIfAny)
                {
                    throw e;
                }
                else
                {
                    clr = System.Drawing.Color.White;
                }
            }
            this.Color.SetColor(clr, 0);
        }

        internal A.GradientStop ToGradientStop()
        {
            A.GradientStop gs = new A.GradientStop();
            if (this.Color.IsRgbColorModelHex) gs.RgbColorModelHex = this.Color.ToRgbColorModelHex();
            else gs.SchemeColor = this.Color.ToSchemeColor();

            gs.Position = Convert.ToInt32(this.Position * 1000m);

            return gs;
        }

        internal SLGradientStop Clone()
        {
            SLGradientStop gs = new SLGradientStop(this.Color.listThemeColors);
            gs.Color = this.Color.Clone();
            gs.decPosition = this.decPosition;

            return gs;
        }
    }
}
