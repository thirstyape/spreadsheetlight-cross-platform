using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for setting a color. This includes using theme colors. This simulates the DocumentFormat.OpenXml.Spreadsheet.Color class.
    /// </summary>
    public class SLColor
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        private System.Drawing.Color clrDisplay = new System.Drawing.Color();
        /// <summary>
        /// The color value.
        /// </summary>
        public System.Drawing.Color Color
        {
            get { return clrDisplay; }
            set
            {
                this.SetAllNull();
                clrDisplay = value;
                this.Rgb = clrDisplay.ToArgb().ToString("X8");
            }
        }

        internal bool? Auto { get; set; }
        internal uint? Indexed { get; set; }
        internal string Rgb { get; set; }
        internal uint? Theme { get; set; }

        internal double? fTint;
        internal double? Tint
        {
            get { return fTint; }
            set
            {
                fTint = value;
                if (fTint != null)
                {
                    if (fTint.Value < -1.0) fTint = -1.0;
                    if (fTint.Value > 1.0) fTint = 1.0;
                }
            }
        }

        internal SLColor(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }
            
            this.listIndexedColors = new List<System.Drawing.Color>();
            for (i = 0; i < IndexedColors.Count; ++i)
            {
                this.listIndexedColors.Add(IndexedColors[i]);
            }

            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.clrDisplay = new System.Drawing.Color();
            this.Auto = null;
            this.Indexed = null;
            this.Rgb = null;
            this.Theme = null;
            this.Tint = null;
        }

        /// <summary>
        /// Whether the color value is empty.
        /// </summary>
        /// <returns>True if the color value is empty. False otherwise.</returns>
        public bool IsEmpty()
        {
            return this.Auto == null && this.Indexed == null && this.Rgb == null && this.Theme == null;
        }

        /// <summary>
        /// Set a color using a theme color.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        public void SetThemeColor(SLThemeColorIndexValues ThemeColorIndex)
        {
            int index = (int)ThemeColorIndex;
            if (index >= 0 && index < this.listThemeColors.Count)
            {
                this.clrDisplay = this.listThemeColors[index];
            }
            this.Auto = null;
            this.Indexed = null;
            this.Rgb = null;
            this.Theme = (uint)ThemeColorIndex;
            this.Tint = null;
        }

        /// <summary>
        /// Set a color using a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetThemeColor(SLThemeColorIndexValues ThemeColorIndex, double Tint)
        {
            System.Drawing.Color clrRgb = new System.Drawing.Color();
            int index = (int)ThemeColorIndex;
            if (index >= 0 && index < this.listThemeColors.Count)
            {
                clrRgb = this.listThemeColors[index];
            }
            this.Auto = null;
            this.Indexed = null;
            this.Rgb = null;
            this.Theme = (uint)ThemeColorIndex;
            this.Tint = Tint;
            this.clrDisplay = SLTool.ToColor(clrRgb, Tint);
        }

        internal BackgroundColor ToBackgroundColor()
        {
            BackgroundColor clr = new BackgroundColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal ForegroundColor ToForegroundColor()
        {
            ForegroundColor clr = new ForegroundColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal TabColor ToTabColor()
        {
            TabColor clr = new TabColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal Color ToSpreadsheetColor()
        {
            Color clr = new Color();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.AxisColor ToAxisColor()
        {
            X14.AxisColor clr = new X14.AxisColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.BarAxisColor ToBarAxisColor()
        {
            X14.BarAxisColor clr = new X14.BarAxisColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.BorderColor ToBorderColor()
        {
            X14.BorderColor clr = new X14.BorderColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.Color ToExcel2010Color()
        {
            X14.Color clr = new X14.Color();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.FillColor ToFillColor()
        {
            X14.FillColor clr = new X14.FillColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.FirstMarkerColor ToFirstMarkerColor()
        {
            X14.FirstMarkerColor clr = new X14.FirstMarkerColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.HighMarkerColor ToHighMarkerColor()
        {
            X14.HighMarkerColor clr = new X14.HighMarkerColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.LastMarkerColor ToLastMarkerColor()
        {
            X14.LastMarkerColor clr = new X14.LastMarkerColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.LowMarkerColor ToLowMarkerColor()
        {
            X14.LowMarkerColor clr = new X14.LowMarkerColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.MarkersColor ToMarkersColor()
        {
            X14.MarkersColor clr = new X14.MarkersColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.NegativeBorderColor ToNegativeBorderColor()
        {
            X14.NegativeBorderColor clr = new X14.NegativeBorderColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.NegativeColor ToNegativeColor()
        {
            X14.NegativeColor clr = new X14.NegativeColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.NegativeFillColor ToNegativeFillColor()
        {
            X14.NegativeFillColor clr = new X14.NegativeFillColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal X14.SeriesColor ToSeriesColor()
        {
            X14.SeriesColor clr = new X14.SeriesColor();
            if (this.Auto != null) clr.Auto = this.Auto.Value;
            if (this.Indexed != null) clr.Indexed = this.Indexed.Value;
            if (this.Rgb != null) clr.Rgb = new HexBinaryValue(this.Rgb);
            if (this.Theme != null) clr.Theme = this.Theme.Value;
            if (this.Tint != null && this.Tint.Value != 0.0) clr.Tint = this.Tint.Value;

            return clr;
        }

        internal void FromBackgroundColor(BackgroundColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromForegroundColor(ForegroundColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromTabColor(TabColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromSpreadsheetColor(Color clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromAxisColor(X14.AxisColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromBarAxisColor(X14.BarAxisColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromBorderColor(X14.BorderColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromExcel2010Color(X14.Color clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromFillColor(X14.FillColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromFirstMarkerColor(X14.FirstMarkerColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromHighMarkerColor(X14.HighMarkerColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromLastMarkerColor(X14.LastMarkerColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromLowMarkerColor(X14.LowMarkerColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromMarkersColor(X14.MarkersColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromNegativeBorderColor(X14.NegativeBorderColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromNegativeColor(X14.NegativeColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromNegativeFillColor(X14.NegativeFillColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        internal void FromSeriesColor(X14.SeriesColor clr)
        {
            this.SetAllNull();
            if (clr.Auto != null) this.Auto = clr.Auto.Value;
            if (clr.Indexed != null) this.Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) this.Rgb = clr.Rgb.Value;
            if (clr.Theme != null) this.Theme = clr.Theme.Value;
            if (clr.Tint != null) this.Tint = clr.Tint.Value;
            this.SetDisplayColor();
        }

        private void SetDisplayColor()
        {
            this.clrDisplay = System.Drawing.Color.FromArgb(255, 0, 0, 0);

            int index = 0;
            if (this.Theme != null)
            {
                index = (int)this.Theme.Value;
                if (index >= 0 && index < this.listThemeColors.Count)
                {
                    this.clrDisplay = System.Drawing.Color.FromArgb(255, this.listThemeColors[index]);
                    if (this.Tint != null)
                    {
                        this.clrDisplay = SLTool.ToColor(this.clrDisplay, this.Tint.Value);
                    }
                }
            }
            else if (this.Rgb != null)
            {
                this.clrDisplay = SLTool.ToColor(this.Rgb);
            }
            else if (this.Indexed != null)
            {
                index = (int)this.Indexed.Value;
                if (index >= 0 && index < this.listIndexedColors.Count)
                {
                    this.clrDisplay = System.Drawing.Color.FromArgb(255, this.listIndexedColors[index]);
                }
            }
        }

        internal void FromHash(string Hash)
        {
            this.SetAllNull();

            string[] sa = Hash.Split(new string[] { SLConstants.XmlColorAttributeSeparator }, StringSplitOptions.None);
            if (sa.Length >= 5)
            {
                if (!sa[0].Equals("null")) this.Auto = bool.Parse(sa[0]);

                if (!sa[1].Equals("null")) this.Indexed = uint.Parse(sa[1]);

                if (!sa[2].Equals("null")) this.Rgb = sa[2];

                if (!sa[3].Equals("null")) this.Theme = uint.Parse(sa[3]);

                if (!sa[4].Equals("null")) this.Tint = double.Parse(sa[4]);
            }

            this.SetDisplayColor();
        }

        internal string ToHash()
        {
            StringBuilder sb = new StringBuilder();

            if (this.Auto != null) sb.AppendFormat("{0}{1}", this.Auto.Value, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            if (this.Indexed != null) sb.AppendFormat("{0}{1}", this.Indexed.Value, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            if (this.Rgb != null) sb.AppendFormat("{0}{1}", this.Rgb, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            if (this.Theme != null) sb.AppendFormat("{0}{1}", this.Theme.Value, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            if (this.Tint != null) sb.AppendFormat("{0}{1}", this.Tint.Value, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            return sb.ToString();
        }

        internal SLColor Clone()
        {
            SLColor clr = new SLColor(this.listThemeColors, this.listIndexedColors);
            clr.clrDisplay = this.clrDisplay;
            clr.Auto = this.Auto;
            clr.Indexed = this.Indexed;
            clr.Rgb = this.Rgb;
            clr.Theme = this.Theme;
            clr.Tint = this.Tint;

            return clr;
        }
    }
}
