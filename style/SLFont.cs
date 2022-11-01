using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for fonts. This simulates the DocumentFormat.OpenXml.Spreadsheet.Font class.
    /// </summary>
    public class SLFont
    {
        internal string MajorFont { get; set; }
        internal string MinorFont { get; set; }
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        /// <summary>
        /// Name of the font.
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// The font character set of the font text. It is recommended not to explicitly set this property. This is used when the given font name is not available on the computer, and a suitable alternative font is used. The character set value is operating system dependent. Possible value (not exhaustive): 0 - ANSI_CHARSET, 1 - DEFAULT_CHARSET, 2 - SYMBOL_CHARSET.
        /// </summary>
        public int? CharacterSet { get; set; }

        /// <summary>
        /// The font family of the font text. It is recommended not to explicitly set this property. Values as follows (might not be exhaustive): 0 - Not applicable, 1 - Roman, 2 - Swiss, 3 - Modern, 4 - Script, 5 - Decorative.
        /// </summary>
        public int? FontFamily { get; set; }

        /// <summary>
        /// Specifies if the font text should be in bold.
        /// </summary>
        public bool? Bold { get; set; }

        /// <summary>
        /// Specifies if the font text should be in italic.
        /// </summary>
        public bool? Italic { get; set; }

        /// <summary>
        /// Specifies if the font text should have a strikethrough.
        /// </summary>
        public bool? Strike { get; set; }

        /// <summary>
        /// Specifies if the inner and outer borders of each character of the font text should be displayed. This makes the font text appear as if in bold.
        /// </summary>
        public bool? Outline { get; set; }

        /// <summary>
        /// Specifies if there's a shadow behind and at the bottom-right of the font text. It is a Macintosh compatibility setting.
        /// It is recommended not to use this property because SpreadsheetML applications are not required to use this property.
        /// </summary>
        public bool? Shadow { get; set; }

        /// <summary>
        /// Specifies if the font text should be squeezed together. It is a Macintosh compatibility setting.
        /// It is recommended not to use this property because SpreadsheetML applications are not required to use this property.
        /// </summary>
        public bool? Condense { get; set; }

        /// <summary>
        /// Specifies if the font text should be stretched out. It is a legacy spreadsheet compatibility setting.
        /// It is recommended not to use this property because SpreadsheetML applications are not required to use this property.
        /// </summary>
        public bool? Extend { get; set; }

        internal bool HasFontColor;
        internal SLColor clrFontColor;
        /// <summary>
        /// The color of the font text.
        /// </summary>
        public System.Drawing.Color FontColor
        {
            get { return clrFontColor.Color; }
            set
            {
                clrFontColor.Color = value;
                HasFontColor = (clrFontColor.Color.IsEmpty) ? false : true;
            }
        }

        /// <summary>
        /// The size of the font text in points (1 point is 1/72 of an inch).
        /// </summary>
        public double? FontSize { get; set; }

        internal bool HasUnderline;
        private UnderlineValues vUnderline;
        // default is single, but for hashing we use none as default
        /// <summary>
        /// Specifies the underline formatting style of the font text.
        /// </summary>
        public UnderlineValues Underline
        {
            get { return vUnderline; }
            set
            {
                vUnderline = value;
                HasUnderline = vUnderline != UnderlineValues.None ? true : false;
            }
        }

        internal bool HasVerticalAlignment;
        private VerticalAlignmentRunValues vVerticalAlignment;
        /// <summary>
        /// Specifies the vertical position of the font text.
        /// </summary>
        public VerticalAlignmentRunValues VerticalAlignment
        {
            get { return vVerticalAlignment; }
            set
            {
                vVerticalAlignment = value;
                HasVerticalAlignment = true;
            }
        }

        internal bool HasFontScheme;
        private FontSchemeValues vFontScheme;
        /// <summary>
        /// Specifies the font scheme. Used particularly as part of a theme definition. A major font scheme is usually used for heading text. A minor font scheme is used for body text.
        /// </summary>
        public FontSchemeValues FontScheme
        {
            get { return vFontScheme; }
            set
            {
                vFontScheme = value;
                HasFontScheme = true;
            }
        }

        /// <summary>
        /// Initializes an instance of SLFont. It is recommended to use CreateFont() of the SLDocument class.
        /// </summary>
        public SLFont()
        {
            this.Initialize(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
        }

        internal SLFont(string MajorFont, string MinorFont, List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            this.Initialize(MajorFont, MinorFont, ThemeColors, IndexedColors);
        }

        private void Initialize(string MajorFont, string MinorFont, List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            this.MajorFont = MajorFont;
            this.MinorFont = MinorFont;

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
            this.FontName = null;
            this.CharacterSet = null;
            this.FontFamily = null;
            this.Bold = null;
            this.Italic = null;
            this.Strike = null;
            this.Outline = null;
            this.Shadow = null;
            this.Condense = null;
            this.Extend = null;
            this.clrFontColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            HasFontColor = false;
            this.FontSize = null;
            this.vUnderline = UnderlineValues.None;
            HasUnderline = false;
            this.vVerticalAlignment = VerticalAlignmentRunValues.Baseline;
            HasVerticalAlignment = false;
            this.vFontScheme = FontSchemeValues.None;
            HasFontScheme = false;
        }

        /// <summary>
        /// Set the font, given a font name and font size.
        /// </summary>
        /// <param name="FontName">The name of the font to be used.</param>
        /// <param name="FontSize">The size of the font in points.</param>
        public void SetFont(string FontName, double FontSize)
        {
            this.FontName = FontName;
            this.FontSize = FontSize;
            this.CharacterSet = null;
            this.FontFamily = null;
            this.vFontScheme = FontSchemeValues.None;
            HasFontScheme = false;
        }

        /// <summary>
        /// Set the font, given a font scheme and font size.
        /// </summary>
        /// <param name="FontScheme">The font scheme. If None is given, the current theme's minor font will be used (but if the theme is changed, the text remains as of the old theme's minor font instead of the new theme's minor font).</param>
        /// <param name="FontSize">The size of the font in points.</param>
        public void SetFont(FontSchemeValues FontScheme, double FontSize)
        {
            switch (FontScheme)
            {
                case FontSchemeValues.Major:
                    this.FontName = this.MajorFont;
                    this.FontScheme = FontSchemeValues.Major;
                    break;
                case FontSchemeValues.Minor:
                    this.FontName = this.MinorFont;
                    this.FontScheme = FontSchemeValues.Minor;
                    break;
                case FontSchemeValues.None:
                    this.FontName = this.MinorFont;
                    this.vFontScheme = FontSchemeValues.None;
                    HasFontScheme = false;
                    break;
            }
            this.FontSize = FontSize;
            this.CharacterSet = null;
            this.FontFamily = null;
        }

        /// <summary>
        /// Set the font color with one of the theme colors.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        public void SetFontThemeColor(SLThemeColorIndexValues ThemeColorIndex)
        {
            this.clrFontColor.SetThemeColor(ThemeColorIndex);
            HasFontColor = (clrFontColor.Color.IsEmpty) ? false : true;
        }

        /// <summary>
        /// Set the font color with one of the theme colors, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetFontThemeColor(SLThemeColorIndexValues ThemeColorIndex, double Tint)
        {
            this.clrFontColor.SetThemeColor(ThemeColorIndex, Tint);
            HasFontColor = (clrFontColor.Color.IsEmpty) ? false : true;
        }

        internal void FromFont(Font f)
        {
            this.SetAllNull();

            if (f.FontName != null && f.FontName.Val != null)
            {
                this.FontName = f.FontName.Val.Value;
            }

            if (f.FontCharSet != null && f.FontCharSet.Val != null)
            {
                this.CharacterSet = f.FontCharSet.Val.Value;
            }

            if (f.FontFamilyNumbering != null && f.FontFamilyNumbering.Val != null)
            {
                this.FontFamily = f.FontFamilyNumbering.Val.Value;
            }

            if (f.Bold != null)
            {
                if (f.Bold.Val == null) this.Bold = true;
                else if (f.Bold.Val.Value) this.Bold = true;
            }

            if (f.Italic != null)
            {
                if (f.Italic.Val == null) this.Italic = true;
                else if (f.Italic.Val.Value) this.Italic = true;
            }

            if (f.Strike != null)
            {
                if (f.Strike.Val == null) this.Strike = true;
                else if (f.Strike.Val.Value) this.Strike = true;
            }

            if (f.Outline != null)
            {
                if (f.Outline.Val == null) this.Outline = true;
                else if (f.Outline.Val.Value) this.Outline = true;
            }

            if (f.Shadow != null)
            {
                if (f.Shadow.Val == null) this.Shadow = true;
                else if (f.Shadow.Val.Value) this.Shadow = true;
            }

            if (f.Condense != null)
            {
                if (f.Condense.Val == null) this.Condense = true;
                else if (f.Condense.Val.Value) this.Condense = true;
            }

            if (f.Extend != null)
            {
                if (f.Extend.Val == null) this.Extend = true;
                else if (f.Extend.Val.Value) this.Extend = true;
            }

            if (f.Color != null)
            {
                this.clrFontColor = new SLColor(this.listThemeColors, this.listIndexedColors);
                this.clrFontColor.FromSpreadsheetColor(f.Color);
                HasFontColor = !this.clrFontColor.IsEmpty();
            }

            if (f.FontSize != null && f.FontSize.Val != null)
            {
                this.FontSize = f.FontSize.Val.Value;
            }

            if (f.Underline != null)
            {
                if (f.Underline.Val != null)
                {
                    this.Underline = f.Underline.Val.Value;
                }
                else
                {
                    this.Underline = UnderlineValues.Single;
                }
            }

            if (f.VerticalTextAlignment != null && f.VerticalTextAlignment.Val != null)
            {
                this.VerticalAlignment = f.VerticalTextAlignment.Val.Value;
            }

            if (f.FontScheme != null && f.FontScheme.Val != null)
            {
                this.FontScheme = f.FontScheme.Val.Value;
            }
        }

        internal Font ToFont()
        {
            Font f = new Font();
            if (this.FontName != null) f.FontName = new FontName() { Val = this.FontName };
            if (this.CharacterSet != null) f.FontCharSet = new FontCharSet() { Val = this.CharacterSet.Value };
            if (this.FontFamily != null) f.FontFamilyNumbering = new FontFamilyNumbering() { Val = this.FontFamily.Value };
            if (this.Bold != null && this.Bold.Value) f.Bold = new Bold();
            if (this.Italic != null && this.Italic.Value) f.Italic = new Italic();
            if (this.Strike != null && this.Strike.Value) f.Strike = new Strike();
            if (this.Outline != null && this.Outline.Value) f.Outline = new Outline();
            if (this.Shadow != null && this.Shadow.Value) f.Shadow = new Shadow();
            if (this.Condense != null && this.Condense.Value) f.Condense = new Condense();
            if (this.Extend != null && this.Extend.Value) f.Extend = new Extend();
            if (HasFontColor) f.Color = this.clrFontColor.ToSpreadsheetColor();
            if (this.FontSize != null) f.FontSize = new FontSize() { Val = this.FontSize.Value };
            if (HasUnderline)
            {
                // default value is Single
                if (this.Underline == UnderlineValues.Single)
                {
                    f.Underline = new Underline();
                }
                else
                {
                    f.Underline = new Underline() { Val = this.Underline };
                }
            }
            if (HasVerticalAlignment) f.VerticalTextAlignment = new VerticalTextAlignment() { Val = this.VerticalAlignment };
            if (HasFontScheme) f.FontScheme = new FontScheme() { Val = this.FontScheme };

            return f;
        }

        internal void FromHash(string Hash)
        {
            Font font = new Font();
            font.InnerXml = Hash;
            this.FromFont(font);
        }

        internal string ToHash()
        {
            Font font = this.ToFont();
            return SLTool.RemoveNamespaceDeclaration(font.InnerXml);
        }

        // SLFont takes on extra duties so you don't have to learn more classes. Just like SLRstType.
        internal A.Paragraph ToParagraph()
        {
            A.Paragraph para = new A.Paragraph();
            para.ParagraphProperties = new A.ParagraphProperties();

            A.DefaultRunProperties defrunprops = new A.DefaultRunProperties();

            string sFont = string.Empty;
            if (this.FontName != null && this.FontName.Length > 0) sFont = this.FontName;

            if (this.HasFontScheme)
            {
                if (this.FontScheme == FontSchemeValues.Major) sFont = "+mj-lt";
                else if (this.FontScheme == FontSchemeValues.Minor) sFont = "+mn-lt";
            }

            if (this.HasFontColor)
            {
                SLA.SLColorTransform clr = new SLA.SLColorTransform(new List<System.Drawing.Color>());
                if (this.clrFontColor.Rgb != null && this.clrFontColor.Rgb.Length > 0)
                {
                    clr.SetColor(SLTool.ToColor(this.clrFontColor.Rgb), 0);

                    defrunprops.Append(new A.SolidFill()
                    {
                        RgbColorModelHex = clr.ToRgbColorModelHex()
                    });
                }
                else if (this.clrFontColor.Theme != null)
                {
                    // potential casting error? If the SLFont class was set properly, there shouldn't be errors...
                    SLThemeColorIndexValues themeindex = (SLThemeColorIndexValues)this.clrFontColor.Theme.Value;
                    if (this.clrFontColor.Tint != null)
                    {
                        clr.SetColor(themeindex, this.clrFontColor.Tint.Value, 0);
                    }
                    else
                    {
                        clr.SetColor(themeindex, 0, 0);
                    }

                    defrunprops.Append(new A.SolidFill()
                    {
                        SchemeColor = clr.ToSchemeColor()
                    });
                }
            }

            if (sFont.Length > 0) defrunprops.Append(new A.LatinFont() { Typeface = sFont });

            if (this.FontSize != null) defrunprops.FontSize = (int)(this.FontSize.Value * 100);

            if (this.Bold != null) defrunprops.Bold = this.Bold.Value;

            if (this.Italic != null) defrunprops.Italic = this.Italic.Value;

            if (this.HasUnderline)
            {
                if (this.Underline == UnderlineValues.Single || this.Underline == UnderlineValues.SingleAccounting)
                {
                    defrunprops.Underline = A.TextUnderlineValues.Single;
                }
                else if (this.Underline == UnderlineValues.Double || this.Underline == UnderlineValues.DoubleAccounting)
                {
                    defrunprops.Underline = A.TextUnderlineValues.Double;
                }
            }

            if (this.Strike != null)
            {
                defrunprops.Strike = this.Strike.Value ? A.TextStrikeValues.SingleStrike : A.TextStrikeValues.NoStrike;
            }

            if (this.HasVerticalAlignment)
            {
                if (this.VerticalAlignment == VerticalAlignmentRunValues.Superscript)
                {
                    defrunprops.Baseline = 30000;
                }
                else if (this.VerticalAlignment == VerticalAlignmentRunValues.Subscript)
                {
                    defrunprops.Baseline = -25000;
                }
                else
                {
                    defrunprops.Baseline = 0;
                }
            }

            para.ParagraphProperties.Append(defrunprops);

            return para;
        }

        /// <summary>
        /// Clone a new instance of SLFont with identical font settings.
        /// </summary>
        /// <returns>An SLFont object with identical font settings.</returns>
        public SLFont Clone()
        {
            SLFont font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);

            font.FontName = this.FontName;
            font.CharacterSet = this.CharacterSet;
            font.FontFamily = this.FontFamily;
            font.Bold = this.Bold;
            font.Italic = this.Italic;
            font.Strike = this.Strike;
            font.Outline = this.Outline;
            font.Shadow = this.Shadow;
            font.Condense = this.Condense;
            font.Extend = this.Extend;
            font.clrFontColor = this.clrFontColor.Clone();
            font.HasFontColor = this.HasFontColor;
            font.FontSize = this.FontSize;
            font.vUnderline = this.vUnderline;
            font.HasUnderline = this.HasUnderline;
            font.vVerticalAlignment = this.vVerticalAlignment;
            font.HasVerticalAlignment = this.HasVerticalAlignment;
            font.vFontScheme = this.vFontScheme;
            font.HasFontScheme = this.HasFontScheme;

            return font;
        }
    }
}
