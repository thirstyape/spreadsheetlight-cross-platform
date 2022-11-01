using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Specifies the named cell style to be used.
    /// </summary>
    public enum SLNamedCellStyleValues
    {
        /// <summary>
        /// Normal
        /// </summary>
        Normal = 0,
        /// <summary>
        /// Bad
        /// </summary>
        Bad,
        /// <summary>
        /// Good
        /// </summary>
        Good,
        /// <summary>
        /// Neutral
        /// </summary>
        Neutral,
        /// <summary>
        /// Calculation
        /// </summary>
        Calculation,
        /// <summary>
        /// Check Cell
        /// </summary>
        CheckCell,
        /// <summary>
        /// Explanatory Text
        /// </summary>
        ExplanatoryText,
        /// <summary>
        /// Input
        /// </summary>
        Input,
        /// <summary>
        /// Linked Cell
        /// </summary>
        LinkedCell,
        /// <summary>
        /// Note
        /// </summary>
        Note,
        /// <summary>
        /// Output
        /// </summary>
        Output,
        /// <summary>
        /// Warning Text
        /// </summary>
        WarningText,
        /// <summary>
        /// Level 1 heading
        /// </summary>
        Heading1,
        /// <summary>
        /// Level 2 heading
        /// </summary>
        Heading2,
        /// <summary>
        /// Level 3 heading
        /// </summary>
        Heading3,
        /// <summary>
        /// Level 4 heading
        /// </summary>
        Heading4,
        /// <summary>
        /// Title
        /// </summary>
        Title,
        /// <summary>
        /// Total
        /// </summary>
        Total,
        /// <summary>
        /// Background color is Accent1 color.
        /// </summary>
        Accent1,
        /// <summary>
        /// Background color is 20% of Accent1 color.
        /// </summary>
        Accent1Percentage20,
        /// <summary>
        /// Background color is 40% of Accent1 color.
        /// </summary>
        Accent1Percentage40,
        /// <summary>
        /// Background color is 60% of Accent1 color.
        /// </summary>
        Accent1Percentage60,
        /// <summary>
        /// Background color is Accent2 color.
        /// </summary>
        Accent2,
        /// <summary>
        /// Background color is 20% of Accent2 color.
        /// </summary>
        Accent2Percentage20,
        /// <summary>
        /// Background color is 40% of Accent2 color.
        /// </summary>
        Accent2Percentage40,
        /// <summary>
        /// Background color is 60% of Accent2 color.
        /// </summary>
        Accent2Percentage60,
        /// <summary>
        /// Background color is Accent3 color.
        /// </summary>
        Accent3,
        /// <summary>
        /// Background color is 20% of Accent3 color.
        /// </summary>
        Accent3Percentage20,
        /// <summary>
        /// Background color is 40% of Accent3 color.
        /// </summary>
        Accent3Percentage40,
        /// <summary>
        /// Background color is 60% of Accent3 color.
        /// </summary>
        Accent3Percentage60,
        /// <summary>
        /// Background color is Accent4 color.
        /// </summary>
        Accent4,
        /// <summary>
        /// Background color is 20% of Accent4 color.
        /// </summary>
        Accent4Percentage20,
        /// <summary>
        /// Background color is 40% of Accent4 color.
        /// </summary>
        Accent4Percentage40,
        /// <summary>
        /// Background color is 60% of Accent4 color.
        /// </summary>
        Accent4Percentage60,
        /// <summary>
        /// Background color is Accent5 color.
        /// </summary>
        Accent5,
        /// <summary>
        /// Background color is 20% of Accent5 color.
        /// </summary>
        Accent5Percentage20,
        /// <summary>
        /// Background color is 40% of Accent5 color.
        /// </summary>
        Accent5Percentage40,
        /// <summary>
        /// Background color is 60% of Accent5 color.
        /// </summary>
        Accent5Percentage60,
        /// <summary>
        /// Background color is Accent6 color.
        /// </summary>
        Accent6,
        /// <summary>
        /// Background color is 20% of Accent6 color.
        /// </summary>
        Accent6Percentage20,
        /// <summary>
        /// Background color is 40% of Accent6 color.
        /// </summary>
        Accent6Percentage40,
        /// <summary>
        /// Background color is 60% of Accent6 color.
        /// </summary>
        Accent6Percentage60,
        /// <summary>
        /// Formats numerical data with a comma as the thousands separator.
        /// </summary>
        Comma,
        /// <summary>
        /// Formats numerical data with a comma as the thousands separator, truncating decimal values.
        /// </summary>
        Comma0,
        /// <summary>
        /// Formats numerical data with a comma as the thousands separator, with $ on the left of the data.
        /// </summary>
        Currency,
        /// <summary>
        /// Formats numerical data with a comma as the thousands separator, with $ on the left of the data, and truncating decimal values.
        /// </summary>
        Currency0,
        /// <summary>
        /// Appends % on the end of the numerical data, and truncating decimal values.
        /// </summary>
        Percentage
    }

    /// <summary>
    /// Encapsulates properties and methods for setting various formatting styles.
    /// </summary>
    public class SLStyle
    {
        internal string MajorFont { get; set; }
        internal string MinorFont { get; set; }
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        internal bool HasAlignment;
        internal SLAlignment alignReal;
        /// <summary>
        /// Specifies the alignment properties for this style.
        /// </summary>
        public SLAlignment Alignment
        {
            get { return alignReal; }
            set
            {
                alignReal = value;
                HasAlignment = true;
            }
        }

        internal bool HasProtection;
        internal SLProtection protectionReal;
        /// <summary>
        /// Specifies the protection properties for this style.
        /// </summary>
        public SLProtection Protection
        {
            get { return protectionReal; }
            set
            {
                protectionReal = value;
                HasProtection = true;
            }
        }

        internal uint? NumberFormatId;
        internal bool HasNumberingFormat;
        internal SLNumberingFormat nfFormatCode;
        /// <summary>
        /// Specifies the (number) format code for this style. Note that the format has to be in invariant-culture mode. So "#,##0.000" is accepted but "#.##0,000" isn't. For cultures with a period as the thousands separator and a comma for the decimal digit separator... sorry.
        /// </summary>
        public string FormatCode
        {
            get { return nfFormatCode.FormatCode; }
            set
            {
                nfFormatCode.FormatCode = value.Trim();
                if (nfFormatCode.FormatCode.Length > 0)
                {
                    HasNumberingFormat = true;
                }
                else
                {
                    HasNumberingFormat = false;
                }
            }
        }

        internal uint? FontId;
        internal bool HasFont;
        internal SLFont fontReal;
        /// <summary>
        /// Specifies the font properties for this style.
        /// </summary>
        public SLFont Font
        {
            get { return fontReal; }
            set
            {
                fontReal = value;
                HasFont = true;
            }
        }

        internal uint? FillId;
        internal bool HasFill;
        internal SLFill fillReal;
        /// <summary>
        /// Specifies the fill properties for this style.
        /// </summary>
        public SLFill Fill
        {
            get { return fillReal; }
            set
            {
                fillReal = value;
                HasFill = true;
            }
        }

        internal uint? BorderId;
        internal bool HasBorder;
        internal SLBorder borderReal;
        /// <summary>
        /// Specifies the border properties for this style.
        /// </summary>
        public SLBorder Border
        {
            get { return borderReal; }
            set
            {
                borderReal = value;
                HasBorder = true;
            }
        }

        // for referencing CellStyles. Not supported yet.
        internal uint? CellStyleFormatId { get; set; }

        /// <summary>
        /// Specifies if the cell content text should be prefixed with a single quotation mark.
        /// </summary>
        public bool? QuotePrefix { get; set; }

        /// <summary>
        /// Specifies if a pivot table dropdown button should be displayed.
        /// </summary>
        public bool? PivotButton { get; set; }

        // for referencing CellStyles. Not supported yet.
        internal bool? ApplyNumberFormat { get; set; }

        // for referencing CellStyles. Not supported yet.
        internal bool? ApplyFont { get; set; }

        // for referencing CellStyles. Not supported yet.
        internal bool? ApplyFill { get; set; }

        // for referencing CellStyles. Not supported yet.
        internal bool? ApplyBorder { get; set; }

        // for referencing CellStyles. Not supported yet.
        internal bool? ApplyAlignment { get; set; }

        // for referencing CellStyles. Not supported yet.
        internal bool? ApplyProtection { get; set; }

        /// <summary>
        /// Initializes an instance of SLStyle. It is recommended to use CreateStyle() of the SLDocument class.
        /// </summary>
        public SLStyle()
        {
            this.Initialize(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
        }

        internal SLStyle(string MajorFont, string MinorFont, List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
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
            RemoveAlignment();
            RemoveProtection();
            RemoveFormatCode();
            RemoveFont();
            RemoveFill();
            RemoveBorder();
            this.CellStyleFormatId = null;
            this.QuotePrefix = null;
            this.PivotButton = null;
            this.ApplyNumberFormat = null;
            this.ApplyFont = null;
            this.ApplyFill = null;
            this.ApplyBorder = null;
            this.ApplyAlignment = null;
            this.ApplyProtection = null;
        }

        /// <summary>
        /// Set the font, given a font name and font size.
        /// </summary>
        /// <param name="FontName">The name of the font to be used.</param>
        /// <param name="FontSize">The size of the font in points.</param>
        public void SetFont(string FontName, double FontSize)
        {
            this.Font.SetFont(FontName, FontSize);
        }

        /// <summary>
        /// Set the font, given a font scheme and font size.
        /// </summary>
        /// <param name="FontScheme">The font scheme. If None is given, the current theme's minor font will be used (but if the theme is changed, the text remains as of the old theme's minor font instead of the new theme's minor font).</param>
        /// <param name="FontSize">The size of the font in points.</param>
        public void SetFont(FontSchemeValues FontScheme, double FontSize)
        {
            this.Font.SetFont(FontScheme, FontSize);
        }

        /// <summary>
        /// Set the font color.
        /// </summary>
        /// <param name="FontColor">The color of the font text.</param>
        public void SetFontColor(System.Drawing.Color FontColor)
        {
            this.Font.FontColor = FontColor;
        }

        /// <summary>
        /// Set the font color with one of the theme colors.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        public void SetFontColor(SLThemeColorIndexValues ThemeColorIndex)
        {
            this.Font.SetFontThemeColor(ThemeColorIndex);
        }

        /// <summary>
        /// Set the font color with one of the theme colors, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetFontColor(SLThemeColorIndexValues ThemeColorIndex, double Tint)
        {
            this.Font.SetFontThemeColor(ThemeColorIndex, Tint);
        }

        /// <summary>
        /// Toggle bold settings.
        /// </summary>
        /// <param name="IsBold">True to set font as bold. False otherwise.</param>
        public void SetFontBold(bool IsBold)
        {
            this.Font.Bold = IsBold;
        }

        /// <summary>
        /// Toggle italic settings.
        /// </summary>
        /// <param name="IsItalic">True to set font as italic. False otherwise.</param>
        public void SetFontItalic(bool IsItalic)
        {
            this.Font.Italic = IsItalic;
        }

        /// <summary>
        /// Set font underline.
        /// </summary>
        /// <param name="FontUnderline">Specifies the underline formatting style of the font text.</param>
        public void SetFontUnderline(UnderlineValues FontUnderline)
        {
            this.Font.Underline = FontUnderline;
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPatternFill(PatternValues PatternType, System.Drawing.Color ForegroundColor, System.Drawing.Color BackgroundColor)
        {
            this.Fill.SetPattern(PatternType, ForegroundColor, BackgroundColor);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPatternFill(PatternValues PatternType, System.Drawing.Color ForegroundColor, SLThemeColorIndexValues BackgroundColorTheme)
        {
            this.Fill.SetPattern(PatternType, ForegroundColor, BackgroundColorTheme);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPatternFill(PatternValues PatternType, System.Drawing.Color ForegroundColor, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            this.Fill.SetPattern(PatternType, ForegroundColor, BackgroundColorTheme, BackgroundColorTint);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPatternFill(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, System.Drawing.Color BackgroundColor)
        {
            this.Fill.SetPattern(PatternType, ForegroundColorTheme, BackgroundColor);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPatternFill(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, SLThemeColorIndexValues BackgroundColorTheme)
        {
            this.Fill.SetPattern(PatternType, ForegroundColorTheme, BackgroundColorTheme);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPatternFill(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            this.Fill.SetPattern(PatternType, ForegroundColorTheme, BackgroundColorTheme, BackgroundColorTint);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPatternFill(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, double ForegroundColorTint, System.Drawing.Color BackgroundColor)
        {
            this.Fill.SetPattern(PatternType, ForegroundColorTheme, ForegroundColorTint, BackgroundColor);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPatternFill(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, double ForegroundColorTint, SLThemeColorIndexValues BackgroundColorTheme)
        {
            this.Fill.SetPattern(PatternType, ForegroundColorTheme, ForegroundColorTint, BackgroundColorTheme);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPatternFill(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, double ForegroundColorTint, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            this.Fill.SetPattern(PatternType, ForegroundColorTheme, ForegroundColorTint, BackgroundColorTheme, BackgroundColorTint);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1">The first color.</param>
        /// <param name="Color2">The second color.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, System.Drawing.Color Color1, System.Drawing.Color Color2)
        {
            this.Fill.SetGradient(ShadingStyle, Color1, Color2);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1">The first color.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, System.Drawing.Color Color1, SLThemeColorIndexValues Color2Theme)
        {
            this.Fill.SetGradient(ShadingStyle, Color1, Color2Theme);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1">The first color.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        /// <param name="Color2Tint">The tint applied to the second theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, System.Drawing.Color Color1, SLThemeColorIndexValues Color2Theme, double Color2Tint)
        {
            this.Fill.SetGradient(ShadingStyle, Color1, Color2Theme, Color2Tint);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color2">The second color.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, System.Drawing.Color Color2)
        {
            this.Fill.SetGradient(ShadingStyle, Color1Theme, Color2);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, SLThemeColorIndexValues Color2Theme)
        {
            this.Fill.SetGradient(ShadingStyle, Color1Theme, Color2Theme);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        /// <param name="Color2Tint">The tint applied to the second theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, SLThemeColorIndexValues Color2Theme, double Color2Tint)
        {
            this.Fill.SetGradient(ShadingStyle, Color1Theme, Color2Theme, Color2Tint);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color1Tint">The tint applied to the first theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Color2">The second color.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, double Color1Tint, System.Drawing.Color Color2)
        {
            this.Fill.SetGradient(ShadingStyle, Color1Theme, Color1Tint, Color2);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color1Tint">The tint applied to the first theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, double Color1Tint, SLThemeColorIndexValues Color2Theme)
        {
            this.Fill.SetGradient(ShadingStyle, Color1Theme, Color1Tint, Color2Theme);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color1Tint">The tint applied to the first theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        /// <param name="Color2Tint">The tint applied to the second theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, double Color1Tint, SLThemeColorIndexValues Color2Theme, double Color2Tint)
        {
            this.Fill.SetGradient(ShadingStyle, Color1Theme, Color1Tint, Color2Theme, Color2Tint);
        }

        /// <summary>
        /// Set the left border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetLeftBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.Border.LeftBorder.BorderStyle = BorderStyle;
            this.Border.LeftBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the left border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetLeftBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.Border.LeftBorder.BorderStyle = BorderStyle;
            this.Border.LeftBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the left border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetLeftBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.Border.LeftBorder.BorderStyle = BorderStyle;
            this.Border.LeftBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the right border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetRightBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.Border.RightBorder.BorderStyle = BorderStyle;
            this.Border.RightBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the right border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetRightBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.Border.RightBorder.BorderStyle = BorderStyle;
            this.Border.RightBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the right border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetRightBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.Border.RightBorder.BorderStyle = BorderStyle;
            this.Border.RightBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the top border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetTopBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.Border.TopBorder.BorderStyle = BorderStyle;
            this.Border.TopBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the top border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetTopBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.Border.TopBorder.BorderStyle = BorderStyle;
            this.Border.TopBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the top border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetTopBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.Border.TopBorder.BorderStyle = BorderStyle;
            this.Border.TopBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the bottom border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetBottomBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.Border.BottomBorder.BorderStyle = BorderStyle;
            this.Border.BottomBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the bottom border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetBottomBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.Border.BottomBorder.BorderStyle = BorderStyle;
            this.Border.BottomBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the bottom border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetBottomBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.Border.BottomBorder.BorderStyle = BorderStyle;
            this.Border.BottomBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the diagonal border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetDiagonalBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.Border.DiagonalBorder.BorderStyle = BorderStyle;
            this.Border.DiagonalBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the diagonal border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetDiagonalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.Border.DiagonalBorder.BorderStyle = BorderStyle;
            this.Border.DiagonalBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the diagonal border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetDiagonalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.Border.DiagonalBorder.BorderStyle = BorderStyle;
            this.Border.DiagonalBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the vertical border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetVerticalBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.Border.VerticalBorder.BorderStyle = BorderStyle;
            this.Border.VerticalBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the vertical border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetVerticalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.Border.VerticalBorder.BorderStyle = BorderStyle;
            this.Border.VerticalBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the vertical border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetVerticalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.Border.VerticalBorder.BorderStyle = BorderStyle;
            this.Border.VerticalBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the horizontal border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetHorizontalBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.Border.HorizontalBorder.BorderStyle = BorderStyle;
            this.Border.HorizontalBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the horizontal border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetHorizontalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.Border.HorizontalBorder.BorderStyle = BorderStyle;
            this.Border.HorizontalBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the horizontal border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetHorizontalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.Border.HorizontalBorder.BorderStyle = BorderStyle;
            this.Border.HorizontalBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Align text vertically.
        /// </summary>
        /// <param name="VerticalAlignment">Specifies the vertical alignment. Default value is Bottom.</param>
        public void SetVerticalAlignment(VerticalAlignmentValues VerticalAlignment)
        {
            this.Alignment.Vertical = VerticalAlignment;
        }

        /// <summary>
        /// Align text horizontally.
        /// </summary>
        /// <param name="HorizontalAlignment">Specifies the horizontal alignment. Default value is General.</param>
        public void SetHorizontalAlignment(HorizontalAlignmentValues HorizontalAlignment)
        {
            this.Alignment.Horizontal = HorizontalAlignment;
        }

        // TODO rotational shortcut functions

        /// <summary>
        /// Toggle wrap text settings.
        /// </summary>
        /// <param name="IsWrapped">True to wrap text. False otherwise.</param>
        public void SetWrapText(bool IsWrapped)
        {
            this.Alignment.WrapText = IsWrapped;
        }

        /// <summary>
        /// Apply a named cell style. Existing styles are kept, unless the chosen named cell style overrides those styles.
        /// </summary>
        /// <param name="NamedCellStyle">The named cell style to be applied.</param>
        public void ApplyNamedCellStyle(SLNamedCellStyleValues NamedCellStyle)
        {
            SLFont font;
            SLFill fill;
            SLBorder border;

            switch (NamedCellStyle)
            {
                case SLNamedCellStyleValues.Normal:
                    RemoveFormatCode();

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    this.Font = font;

                    RemoveFill();
                    RemoveBorder();

                    // normal is the only one that removes alignment
                    RemoveAlignment();
                    break;
                case SLNamedCellStyleValues.Bad:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0x9C, 0, 0x06);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xC7, 0xCE));
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Good:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0, 0x61, 0);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xFF, 0xC6, 0xEF, 0xCE));
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Neutral:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0x9C, 0x65, 0);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEB, 0x9C));
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Calculation:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.Bold = true;
                    font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0xFA, 0x7D, 0);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xFF, 0xF2, 0xF2, 0xF2));
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.LeftBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x7F, 0x7F, 0x7F);
                    border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
                    border.RightBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x7F, 0x7F, 0x7F);
                    border.RightBorder.BorderStyle = BorderStyleValues.Thin;
                    border.TopBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x7F, 0x7F, 0x7F);
                    border.TopBorder.BorderStyle = BorderStyleValues.Thin;
                    border.BottomBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x7F, 0x7F, 0x7F);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.CheckCell:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.Bold = true;
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xFF, 0xA5, 0xA5, 0xA5));
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.LeftBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x3F);
                    border.LeftBorder.BorderStyle = BorderStyleValues.Double;
                    border.RightBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x3F);
                    border.RightBorder.BorderStyle = BorderStyleValues.Double;
                    border.TopBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x3F);
                    border.TopBorder.BorderStyle = BorderStyleValues.Double;
                    border.BottomBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x3F);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Double;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.ExplanatoryText:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.Italic = true;
                    font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0x7F, 0x7F, 0x7F);
                    this.Font = font;

                    // no change to fill

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Input:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x76);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xCC, 0x99));
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.LeftBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x7F, 0x7F, 0x7F);
                    border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
                    border.RightBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x7F, 0x7F, 0x7F);
                    border.RightBorder.BorderStyle = BorderStyleValues.Thin;
                    border.TopBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x7F, 0x7F, 0x7F);
                    border.TopBorder.BorderStyle = BorderStyleValues.Thin;
                    border.BottomBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x7F, 0x7F, 0x7F);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.LinkedCell:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0xFA, 0x7D, 0);
                    this.Font = font;

                    // no change to fill

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.BottomBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x80, 0x01);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Double;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.Note:
                    // no change to format code

                    // Note doesn't change font or font size

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xFF, 0xCC));
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.LeftBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0xB2, 0xB2, 0xB2);
                    border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
                    border.RightBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0xB2, 0xB2, 0xB2);
                    border.RightBorder.BorderStyle = BorderStyleValues.Thin;
                    border.TopBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0xB2, 0xB2, 0xB2);
                    border.TopBorder.BorderStyle = BorderStyleValues.Thin;
                    border.BottomBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0xB2, 0xB2, 0xB2);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.Output:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.Bold = true;
                    font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x3F);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(0xFF, 0xF2, 0xF2, 0xF2));
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.LeftBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x3F);
                    border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
                    border.RightBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x3F);
                    border.RightBorder.BorderStyle = BorderStyleValues.Thin;
                    border.TopBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x3F);
                    border.TopBorder.BorderStyle = BorderStyleValues.Thin;
                    border.BottomBorder.Color = System.Drawing.Color.FromArgb(0xFF, 0x3F, 0x3F, 0x3F);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.WarningText:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0, 0);
                    this.Font = font;

                    // no change to fill

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Heading1:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.Heading1FontSize);
                    font.Bold = true;
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark2Color);
                    this.Font = font;

                    // no change to fill

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.BottomBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent1Color);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Thick;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.Heading2:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.Heading2FontSize);
                    font.Bold = true;
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark2Color);
                    this.Font = font;

                    // no change to fill

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.BottomBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent1Color, 0.499984740745262);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Thick;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.Heading3:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.Bold = true;
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark2Color);
                    this.Font = font;

                    // no change to fill

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.BottomBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent1Color, 0.399975585192419);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Medium;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.Heading4:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.Bold = true;
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark2Color);
                    this.Font = font;

                    // no change to fill

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Title:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Major, SLConstants.TitleFontSize);
                    font.Bold = true;
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark2Color);
                    this.Font = font;

                    // no change to fill

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Total:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.Bold = true;
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    // no change to fill

                    border = new SLBorder(this.listThemeColors, this.listIndexedColors);
                    border.TopBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent1Color);
                    border.TopBorder.BorderStyle = BorderStyleValues.Thin;
                    border.BottomBorder.SetBorderThemeColor(SLThemeColorIndexValues.Accent1Color);
                    border.BottomBorder.BorderStyle = BorderStyleValues.Double;
                    this.Border = border;
                    break;
                case SLNamedCellStyleValues.Accent1:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent1Color);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent1Percentage20:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent1Color, 0.799981688894314);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent1Percentage40:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent1Color, 0.599993896298105);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent1Percentage60:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent1Color, 0.399975585192419);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent2:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent2Color);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent2Percentage20:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent2Color, 0.799981688894314);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent2Percentage40:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent2Color, 0.599993896298105);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent2Percentage60:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent2Color, 0.399975585192419);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent3:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent3Color);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent3Percentage20:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent3Color, 0.799981688894314);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent3Percentage40:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent3Color, 0.599993896298105);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent3Percentage60:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent3Color, 0.399975585192419);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent4:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent4Color);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent4Percentage20:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent4Color, 0.799981688894314);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent4Percentage40:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent4Color, 0.599993896298105);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent4Percentage60:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent4Color, 0.399975585192419);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent5:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent5Color);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent5Percentage20:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent5Color, 0.799981688894314);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent5Percentage40:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent5Color, 0.599993896298105);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent5Percentage60:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent5Color, 0.399975585192419);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent6:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent6Color);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent6Percentage20:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent6Color, 0.799981688894314);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent6Percentage40:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent6Color, 0.599993896298105);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Accent6Percentage60:
                    // no change to format code

                    font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
                    font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    font.SetFontThemeColor(SLThemeColorIndexValues.Light1Color);
                    this.Font = font;

                    fill = new SLFill(this.listThemeColors, this.listIndexedColors);
                    fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent6Color, 0.399975585192419);
                    fill.SetPatternType(PatternValues.Solid);
                    this.Fill = fill;

                    // no change to border
                    break;
                case SLNamedCellStyleValues.Comma:
                    this.FormatCode = "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)";
                    // not using the "builtin" format ID
                    //this.nfFormatCode.NumberFormatId = 43;

                    // the font, fill and border are not changed

                    // TODO get "actual" comma character from regional settings?
                    break;
                case SLNamedCellStyleValues.Comma0:
                    this.FormatCode = "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)";
                    // not using the "builtin" format ID
                    //this.nfFormatCode.NumberFormatId = 41;

                    // the font, fill and border are not changed

                    // TODO get "actual" comma character from regional settings?
                    break;
                case SLNamedCellStyleValues.Currency:
                    this.FormatCode = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)";
                    // not using the "builtin" format ID
                    //this.nfFormatCode.NumberFormatId = 44;

                    // the font, fill and border are not changed

                    // TODO get "actual" currency character from regional settings?
                    break;
                case SLNamedCellStyleValues.Currency0:
                    this.FormatCode = "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)";
                    // not using the "builtin" format ID
                    //this.nfFormatCode.NumberFormatId = 42;

                    // the font, fill and border are not changed

                    // TODO get "actual" currency character from regional settings?
                    break;
                case SLNamedCellStyleValues.Percentage:
                    this.FormatCode = "0%";
                    // not using the "builtin" format ID
                    //this.nfFormatCode.NumberFormatId = 9;

                    // the font, fill and border are not changed
                    break;
            }
        }

        /// <summary>
        /// Remove any existing alignment properties.
        /// </summary>
        public void RemoveAlignment()
        {
            this.alignReal = new SLAlignment();
            HasAlignment = false;
        }

        /// <summary>
        /// Remove any existing protection properties.
        /// </summary>
        public void RemoveProtection()
        {
            this.protectionReal = new SLProtection();
            HasProtection = false;
        }

        /// <summary>
        /// Remove any existing format code.
        /// </summary>
        public void RemoveFormatCode()
        {
            this.NumberFormatId = null;
            this.nfFormatCode = new SLNumberingFormat();
            HasNumberingFormat = false;
        }

        /// <summary>
        /// Remove any existing font properties.
        /// </summary>
        public void RemoveFont()
        {
            this.FontId = null;
            this.fontReal = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
            HasFont = false;
        }

        /// <summary>
        /// Remove any existing fill properties.
        /// </summary>
        public void RemoveFill()
        {
            this.FillId = null;
            this.fillReal = new SLFill(this.listThemeColors, this.listIndexedColors);
            HasFill = false;
        }

        /// <summary>
        /// Remove any existing border properties.
        /// </summary>
        public void RemoveBorder()
        {
            this.BorderId = null;
            this.borderReal = new SLBorder(this.listThemeColors, this.listIndexedColors);
            HasBorder = false;
        }

        internal void MergeStyle(SLStyle NewStyle)
        {
            NewStyle.Sync();

            if (NewStyle.HasAlignment) this.Alignment = NewStyle.Alignment.Clone();
            if (NewStyle.HasProtection) this.Protection = NewStyle.Protection.Clone();

            // if the new style format code is non-empty, then overwrite (even if new one is General)
            if (!string.IsNullOrEmpty(NewStyle.FormatCode)) this.FormatCode = NewStyle.FormatCode;

            if (NewStyle.HasFont)
            {
                // what's the point if there's no font name?
                this.fontReal.FontName = NewStyle.fontReal.FontName;

                if (NewStyle.fontReal.CharacterSet != null) this.fontReal.CharacterSet = NewStyle.fontReal.CharacterSet.Value;
                if (NewStyle.fontReal.FontFamily != null) this.fontReal.FontFamily = NewStyle.fontReal.FontFamily.Value;
                if (NewStyle.fontReal.Bold != null) this.fontReal.Bold = NewStyle.fontReal.Bold.Value;
                if (NewStyle.fontReal.Italic != null) this.fontReal.Italic = NewStyle.fontReal.Italic.Value;
                if (NewStyle.fontReal.Strike != null) this.fontReal.Strike = NewStyle.fontReal.Strike.Value;
                if (NewStyle.fontReal.Outline != null) this.fontReal.Outline = NewStyle.fontReal.Outline.Value;
                if (NewStyle.fontReal.Shadow != null) this.fontReal.Shadow = NewStyle.fontReal.Shadow.Value;
                if (NewStyle.fontReal.Condense != null) this.fontReal.Condense = NewStyle.fontReal.Condense.Value;
                if (NewStyle.fontReal.Extend != null) this.fontReal.Extend = NewStyle.fontReal.Extend.Value;
                if (NewStyle.fontReal.HasFontColor)
                {
                    this.fontReal.clrFontColor = NewStyle.fontReal.clrFontColor.Clone();
                    this.fontReal.HasFontColor = (this.fontReal.clrFontColor.Color.IsEmpty) ? false : true;
                }
                
                // no point if there's no font size either
                this.fontReal.FontSize = NewStyle.fontReal.FontSize;
                
                if (NewStyle.fontReal.HasUnderline) this.fontReal.Underline = NewStyle.fontReal.Underline;
                if (NewStyle.fontReal.HasVerticalAlignment) this.fontReal.VerticalAlignment = NewStyle.fontReal.VerticalAlignment;

                if (NewStyle.fontReal.HasFontScheme)
                {
                    this.fontReal.FontScheme = NewStyle.fontReal.FontScheme;
                }
                else
                {
                    this.fontReal.FontScheme = FontSchemeValues.None;
                    this.fontReal.HasFontScheme = false;
                }

                this.HasFont = true;
            }

            if (NewStyle.HasFill) this.Fill = NewStyle.Fill.Clone();
            
            if (NewStyle.HasBorder)
            {
                if (NewStyle.borderReal.HasLeftBorder) this.borderReal.LeftBorder = NewStyle.borderReal.LeftBorder.Clone();
                if (NewStyle.borderReal.HasRightBorder) this.borderReal.RightBorder = NewStyle.borderReal.RightBorder.Clone();
                if (NewStyle.borderReal.HasTopBorder) this.borderReal.TopBorder = NewStyle.borderReal.TopBorder.Clone();
                if (NewStyle.borderReal.HasBottomBorder) this.borderReal.BottomBorder = NewStyle.borderReal.BottomBorder.Clone();
                if (NewStyle.borderReal.HasDiagonalBorder) this.borderReal.DiagonalBorder = NewStyle.borderReal.DiagonalBorder.Clone();
                if (NewStyle.borderReal.HasVerticalBorder) this.borderReal.VerticalBorder = NewStyle.borderReal.VerticalBorder.Clone();
                if (NewStyle.borderReal.HasHorizontalBorder) this.borderReal.HorizontalBorder = NewStyle.borderReal.HorizontalBorder.Clone();
                this.borderReal.DiagonalUp = NewStyle.borderReal.DiagonalUp;
                this.borderReal.DiagonalDown = NewStyle.borderReal.DiagonalDown;
                this.borderReal.Outline = NewStyle.borderReal.Outline;
                
                this.HasBorder = true;
            }
        }

        internal void Sync()
        {
            HasAlignment = Alignment.HasHorizontal || Alignment.HasVertical || Alignment.TextRotation != null || Alignment.WrapText != null || Alignment.Indent != null || Alignment.RelativeIndent != null || Alignment.JustifyLastLine != null || Alignment.ShrinkToFit != null || Alignment.HasReadingOrder;
            HasProtection = Protection.Locked != null || Protection.Hidden != null;
            //HasNumberingFormat
            HasFont = Font.FontName != null || Font.CharacterSet != null || Font.FontFamily != null || Font.Bold != null || Font.Italic != null || Font.Strike != null || Font.Outline != null || Font.Shadow != null || Font.Condense != null || Font.Extend != null || Font.HasFontColor || Font.FontSize != null || Font.HasUnderline || Font.HasVerticalAlignment || Font.HasFontScheme;
            HasFill = Fill.HasBeenAssignedValues;
            Border.Sync();
            HasBorder = Border.HasLeftBorder || Border.HasRightBorder || Border.HasTopBorder || Border.HasBottomBorder || Border.HasDiagonalBorder || Border.HasVerticalBorder || Border.HasHorizontalBorder || Border.DiagonalUp != null || Border.DiagonalDown != null || Border.Outline != null;
        }

        internal void FromCellFormat(CellFormat cf)
        {
            this.SetAllNull();

            if (cf.Alignment != null)
            {
                HasAlignment = true;
                this.alignReal = new SLAlignment();
                this.alignReal.FromAlignment(cf.Alignment);
            }

            if (cf.Protection != null)
            {
                HasProtection = true;
                this.protectionReal = new SLProtection();
                this.protectionReal.FromProtection(cf.Protection);
            }

            if (cf.NumberFormatId != null) this.NumberFormatId = cf.NumberFormatId.Value;

            if (cf.FontId != null) this.FontId = cf.FontId.Value;

            if (cf.FillId != null) this.FillId = cf.FillId.Value;

            if (cf.BorderId != null) this.BorderId = cf.BorderId.Value;

            if (cf.FormatId != null) this.CellStyleFormatId = cf.FormatId.Value;

            if (cf.QuotePrefix != null) this.QuotePrefix = cf.QuotePrefix.Value;

            if (cf.PivotButton != null) this.PivotButton = cf.PivotButton.Value;

            if (cf.ApplyNumberFormat != null) this.ApplyNumberFormat = cf.ApplyNumberFormat.Value;

            if (cf.ApplyFont != null) this.ApplyFont = cf.ApplyFont.Value;

            if (cf.ApplyFill != null) this.ApplyFill = cf.ApplyFill.Value;

            if (cf.ApplyBorder != null) this.ApplyBorder = cf.ApplyBorder.Value;

            if (cf.ApplyAlignment != null) this.ApplyAlignment = cf.ApplyAlignment.Value;

            if (cf.ApplyProtection != null) this.ApplyProtection = cf.ApplyProtection.Value;

            Sync();
        }

        /// <summary>
        /// IMPORTANT! Fill the indices for numbering format, font, fill and border!
        /// </summary>
        /// <returns></returns>
        internal CellFormat ToCellFormat()
        {
            Sync();

            CellFormat cf = new CellFormat();
            if (HasAlignment) cf.Alignment = this.Alignment.ToAlignment();
            if (HasProtection) cf.Protection = this.Protection.ToProtection();

            if (this.NumberFormatId != null) cf.NumberFormatId = this.NumberFormatId.Value;
            if (this.FontId != null) cf.FontId = this.FontId.Value;
            if (this.FillId != null) cf.FillId = this.FillId.Value;
            if (this.BorderId != null) cf.BorderId = this.BorderId.Value;

            if (this.CellStyleFormatId != null) cf.FormatId = this.CellStyleFormatId.Value;
            if (this.QuotePrefix != null && this.QuotePrefix.Value) cf.QuotePrefix = this.QuotePrefix.Value;
            if (this.PivotButton != null && this.PivotButton.Value) cf.PivotButton = this.PivotButton.Value;
            if (this.ApplyNumberFormat != null) cf.ApplyNumberFormat = this.ApplyNumberFormat.Value;
            if (this.ApplyFont != null) cf.ApplyFont = this.ApplyFont.Value;
            if (this.ApplyFill != null) cf.ApplyFill = this.ApplyFill.Value;
            if (this.ApplyBorder != null) cf.ApplyBorder = this.ApplyBorder.Value;
            if (this.ApplyAlignment != null) cf.ApplyAlignment = this.ApplyAlignment.Value;
            if (this.ApplyProtection != null) cf.ApplyProtection = this.ApplyProtection.Value;

            return cf;
        }

        internal void FromHash(string Hash)
        {
            this.SetAllNull();

            string[] saElementAttribute = Hash.Split(new string[] { SLConstants.XmlStyleElementAttributeSeparator }, StringSplitOptions.None);

            if (saElementAttribute.Length >= 7)
            {
                if (!saElementAttribute[0].Equals("null")) this.alignReal.FromHash(saElementAttribute[0]);

                if (!saElementAttribute[1].Equals("null")) this.protectionReal.FromHash(saElementAttribute[1]);

                if (!saElementAttribute[2].Equals("null")) this.nfFormatCode.FromHash(saElementAttribute[2]);

                if (!saElementAttribute[3].Equals("null")) this.fontReal.FromHash(saElementAttribute[3]);

                if (!saElementAttribute[4].Equals("null")) this.fillReal.FromHash(saElementAttribute[4]);

                if (!saElementAttribute[5].Equals("null")) this.borderReal.FromHash(saElementAttribute[5]);

                string[] sa = saElementAttribute[6].Split(new string[] { SLConstants.XmlStyleAttributeSeparator }, StringSplitOptions.None);
                if (sa.Length >= 9)
                {
                    if (!sa[0].Equals("null")) this.CellStyleFormatId = uint.Parse(sa[0]);

                    if (!sa[1].Equals("null")) this.QuotePrefix = bool.Parse(sa[1]);

                    if (!sa[2].Equals("null")) this.PivotButton = bool.Parse(sa[2]);

                    if (!sa[3].Equals("null")) this.ApplyNumberFormat = bool.Parse(sa[3]);

                    if (!sa[4].Equals("null")) this.ApplyFont = bool.Parse(sa[4]);

                    if (!sa[5].Equals("null")) this.ApplyFill = bool.Parse(sa[5]);

                    if (!sa[6].Equals("null")) this.ApplyBorder = bool.Parse(sa[6]);

                    if (!sa[7].Equals("null")) this.ApplyAlignment = bool.Parse(sa[7]);

                    if (!sa[8].Equals("null")) this.ApplyProtection = bool.Parse(sa[8]);
                }
            }

            Sync();
        }

        internal string ToHash()
        {
            Sync();

            StringBuilder sb = new StringBuilder();

            if (HasAlignment) sb.AppendFormat("{0}{1}", this.alignReal.ToHash(), SLConstants.XmlStyleElementAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleElementAttributeSeparator);

            if (HasProtection) sb.AppendFormat("{0}{1}", this.protectionReal.ToHash(), SLConstants.XmlStyleElementAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleElementAttributeSeparator);

            if (this.nfFormatCode.FormatCode.Length > 0) sb.AppendFormat("{0}{1}", this.nfFormatCode.FormatCode, SLConstants.XmlStyleElementAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleElementAttributeSeparator);

            if (HasFont) sb.AppendFormat("{0}{1}", this.fontReal.ToHash(), SLConstants.XmlStyleElementAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleElementAttributeSeparator);

            if (HasFill) sb.AppendFormat("{0}{1}", this.fillReal.ToHash(), SLConstants.XmlStyleElementAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleElementAttributeSeparator);

            if (HasBorder) sb.AppendFormat("{0}{1}", this.borderReal.ToHash(), SLConstants.XmlStyleElementAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleElementAttributeSeparator);

            if (this.CellStyleFormatId != null) sb.AppendFormat("{0}{1}", this.CellStyleFormatId.Value.ToString(CultureInfo.InvariantCulture), SLConstants.XmlStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleAttributeSeparator);

            if (this.QuotePrefix != null) sb.AppendFormat("{0}{1}", this.QuotePrefix.Value, SLConstants.XmlStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleAttributeSeparator);

            if (this.PivotButton != null) sb.AppendFormat("{0}{1}", this.PivotButton.Value, SLConstants.XmlStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleAttributeSeparator);

            if (this.ApplyNumberFormat != null) sb.AppendFormat("{0}{1}", this.ApplyNumberFormat.Value, SLConstants.XmlStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleAttributeSeparator);

            if (this.ApplyFont != null) sb.AppendFormat("{0}{1}", this.ApplyFont.Value, SLConstants.XmlStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleAttributeSeparator);

            if (this.ApplyFill != null) sb.AppendFormat("{0}{1}", this.ApplyFill.Value, SLConstants.XmlStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleAttributeSeparator);

            if (this.ApplyBorder != null) sb.AppendFormat("{0}{1}", this.ApplyBorder.Value, SLConstants.XmlStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleAttributeSeparator);

            if (this.ApplyAlignment != null) sb.AppendFormat("{0}{1}", this.ApplyAlignment.Value, SLConstants.XmlStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleAttributeSeparator);

            if (this.ApplyProtection != null) sb.AppendFormat("{0}{1}", this.ApplyProtection.Value, SLConstants.XmlStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlStyleAttributeSeparator);

            return sb.ToString();
        }

        internal string WriteToXmlTag()
        {
            Sync();

            StringBuilder sb = new StringBuilder();
            sb.Append("<x:xf");

            //if (this.nfFormatCode.FormatCode.Length > 0) sb.AppendFormat(" numFmtId=\"{0}\"", this.nfFormatCode.SaveToStylesheet());

            //if (HasFont) sb.AppendFormat(" fontId=\"{0}\"", this.Font.SaveToStylesheet());
            //if (HasFill) sb.AppendFormat(" fillId=\"{0}\"", this.Fill.SaveToStylesheet());
            //if (HasBorder) sb.AppendFormat(" borderId=\"{0}\"", this.Border.SaveToStylesheet());

            if (this.NumberFormatId != null) sb.AppendFormat(" numFmtId=\"{0}\"", this.NumberFormatId.Value.ToString(CultureInfo.InvariantCulture));
            if (this.FontId != null) sb.AppendFormat(" fontId=\"{0}\"", this.FontId.Value.ToString(CultureInfo.InvariantCulture));
            if (this.FillId != null) sb.AppendFormat(" fillId=\"{0}\"", this.FillId.Value.ToString(CultureInfo.InvariantCulture));
            if (this.BorderId != null) sb.AppendFormat(" borderId=\"{0}\"", this.BorderId.Value.ToString(CultureInfo.InvariantCulture));

            if (this.CellStyleFormatId != null) sb.AppendFormat(" xfId=\"{0}\"", this.CellStyleFormatId.Value);
            if (this.QuotePrefix != null && this.QuotePrefix.Value) sb.Append(" quotePrefix=\"1\"");
            if (this.PivotButton != null && this.PivotButton.Value) sb.Append(" pivotButton=\"1\"");
            if (this.ApplyNumberFormat != null) sb.AppendFormat(" applyNumberFormat=\"{0}\"", this.ApplyNumberFormat.Value ? "1" : "0");
            if (this.ApplyFont != null) sb.AppendFormat(" applyFont=\"{0}\"", this.ApplyFont.Value ? "1" : "0");
            if (this.ApplyFill != null) sb.AppendFormat(" applyFill=\"{0}\"", this.ApplyFill.Value ? "1" : "0");
            if (this.ApplyBorder != null) sb.AppendFormat(" applyBorder=\"{0}\"", this.ApplyBorder.Value ? "1" : "0");
            if (this.ApplyAlignment != null) sb.AppendFormat(" applyAlignment=\"{0}\"", this.ApplyAlignment.Value ? "1" : "0");
            if (this.ApplyProtection != null) sb.AppendFormat(" applyProtection=\"{0}\"", this.ApplyProtection.Value ? "1" : "0");

            if (this.HasAlignment || this.HasProtection)
            {
                sb.Append(">");
                if (this.HasAlignment)
                {
                    sb.Append(this.alignReal.WriteToXmlTag());
                }
                if (this.HasProtection)
                {
                    sb.Append(this.protectionReal.WriteToXmlTag());
                }
                sb.Append("</x:xf>");
            }
            else
            {
                sb.Append(" />");
            }

            return sb.ToString();
        }

        /// <summary>
        /// Clone a new instance of SLStyle with identical style settings.
        /// </summary>
        /// <returns>An SLStyle object with identical style settings.</returns>
        public SLStyle Clone()
        {
            SLStyle style = new SLStyle(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
            style.HasAlignment = this.HasAlignment;
            style.alignReal = this.alignReal.Clone();
            style.HasProtection = this.HasProtection;
            style.protectionReal = this.protectionReal.Clone();
            style.NumberFormatId = this.NumberFormatId;
            style.HasNumberingFormat = this.HasNumberingFormat;
            style.nfFormatCode = this.nfFormatCode.Clone();
            style.FontId = this.FontId;
            style.HasFont = this.HasFont;
            style.fontReal = this.fontReal.Clone();
            style.FillId = this.FillId;
            style.HasFill = this.HasFill;
            style.fillReal = this.fillReal.Clone();
            style.BorderId = this.BorderId;
            style.HasBorder = this.HasBorder;
            style.borderReal = this.borderReal.Clone();
            style.CellStyleFormatId = this.CellStyleFormatId;
            style.QuotePrefix = this.QuotePrefix;
            style.PivotButton = this.PivotButton;
            style.ApplyNumberFormat = this.ApplyNumberFormat;
            style.ApplyFont = this.ApplyFont;
            style.ApplyFill = this.ApplyFill;
            style.ApplyBorder = this.ApplyBorder;
            style.ApplyAlignment = this.ApplyAlignment;
            style.ApplyProtection = this.ApplyProtection;

            return style;
        }
    }
}
