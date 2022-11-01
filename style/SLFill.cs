using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Specifies gradient shading style options.
    /// </summary>
    public enum SLGradientShadingStyleValues
    {
        /// <summary>
        /// Gradient with color 1 at the top to color 2 at the bottom.
        /// </summary>
        Horizontal1 = 0,
        /// <summary>
        /// Gradient with color 2 at the top to color 1 at the bottom.
        /// </summary>
        Horizontal2,
        /// <summary>
        /// Gradient with color 1 at the top to color 2 in the middle, to color 1 at the bottom.
        /// </summary>
        Horizontal3,
        /// <summary>
        /// Gradient with color 1 on the left to color 2 on the right.
        /// </summary>
        Vertical1,
        /// <summary>
        /// Gradient with color 2 on the left to color 1 on the right.
        /// </summary>
        Vertical2,
        /// <summary>
        /// Gradient with color 1 on the left to color 2 in the middle, to color 1 on the right.
        /// </summary>
        Vertical3,
        /// <summary>
        /// Gradient with color 1 at top-left corner to color 2 at bottom-right corner.
        /// </summary>
        DiagonalUp1,
        /// <summary>
        /// Gradient with color 2 at top-left corner to color 1 at the bottom-right corner.
        /// </summary>
        DiagonalUp2,
        /// <summary>
        /// Gradient with color 1 at top-left corner to color 2 in the middle, to color 1 at the bottom-right corner.
        /// </summary>
        DiagonalUp3,
        /// <summary>
        /// Gradient with color 1 at the top-right corner to color 2 at the bottom-left corner.
        /// </summary>
        DiagonalDown1,
        /// <summary>
        /// Gradient with color 2 at the top-right corner to color 1 at the bottom-left corner.
        /// </summary>
        DiagonalDown2,
        /// <summary>
        /// Gradient with color 1 at the top-right corner to color 2 in the middle, to color 1 at the bottom-left corner.
        /// </summary>
        DiagonalDown3,
        /// <summary>
        /// Gradient with color 1 at the top-left corner, and color 2 at the top-right, bottom-left and bottom-right corners.
        /// </summary>
        Corner1,
        /// <summary>
        /// Gradient with color 1 at the top-right corner, and color 2 at the top-left, bottom-left and bottom-right corners.
        /// </summary>
        Corner2,
        /// <summary>
        /// Gradient with color 1 at the bottom-left corner, and color 2 at the top-left, top-right and bottom-right corners.
        /// </summary>
        Corner3,
        /// <summary>
        /// Gradient with color 1 at the bottom-right corner, and color 2 at the top-left, top-right and bottom-left corners.
        /// </summary>
        Corner4,
        /// <summary>
        /// Gradient with color 1 in the center, and color 2 at the four corners.
        /// </summary>
        FromCenter
    }

    /// <summary>
    /// Encapsulates properties and methods for specifying fill options, such as foreground and background colors.
    /// This simulates the DocumentFormat.OpenXml.Spreadsheet.Fill class.
    /// </summary>
    public class SLFill
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        // this is for the parent class's (SLStyle)
        // equivalent of "HasFill" boolean 
        internal bool HasBeenAssignedValues;

        // as opposed to using GradientFill
        internal bool UsePatternFill;

        private SLPatternFill pfReal;
        private SLGradientFill gfReal;

        /// <summary>
        /// Color of the foreground. This is read-only. Use one of the methods to set the color.
        /// </summary>
        public System.Drawing.Color PatternForegroundColor
        {
            get { return pfReal.ForegroundColor; }
        }

        /// <summary>
        /// Color of the background. This is read-only. Use one of the methods to set the color.
        /// </summary>
        public System.Drawing.Color PatternBackgroundColor
        {
            get { return pfReal.BackgroundColor; }
        }

        /// <summary>
        /// Pattern type of the fill. This is read-only. Use one of the methods to set the pattern type.
        /// </summary>
        public PatternValues PatternType
        {
            get
            {
                if (pfReal.HasPatternType) return pfReal.PatternType;
                else return PatternValues.None;
            }
        }

        /// <summary>
        /// Gradient type of the fill. This is read-only. Use one of the methods to set the gradient type.
        /// </summary>
        public GradientValues GradientType
        {
            get
            {
                if (gfReal.HasType) return gfReal.Type;
                else return GradientValues.Linear;
            }
        }

        /// <summary>
        /// The angle in the direction from which the first color starts. The end color is at 180 degrees from this angle. 0 degrees means start from left, 90 degrees from the top, 180 degrees from the right and 270 degrees from below.
        /// This is read-only. Use one of the methods to set the angle.
        /// </summary>
        public double GradientDegree
        {
            get
            {
                if (gfReal.Degree != null) return gfReal.Degree.Value;
                else return 0.0;
            }
        }

        /// <summary>
        /// Specifies position of the first color at the left edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the left edge of the cell, and 1.0 means it's on the right edge.
        /// This is read-only. Use one of the methods to set the position.
        /// </summary>
        public double GradientLeft
        {
            get
            {
                if (gfReal.Left != null) return gfReal.Left.Value;
                else return 0.0;
            }
        }

        /// <summary>
        /// Specifies position of the first color at the right edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the left edge of the cell, and 1.0 means it's on the right edge.
        /// This is read-only. Use one of the methods to set the position.
        /// </summary>
        public double GradientRight
        {
            get
            {
                if (gfReal.Right != null) return gfReal.Right.Value;
                else return 0.0;
            }
        }

        /// <summary>
        /// Specifies position of the first color at the top edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the top edge of the cell, and 1.0 means it's on the bottom edge.
        /// This is read-only. Use one of the methods to set the position.
        /// </summary>
        public double GradientTop
        {
            get
            {
                if (gfReal.Top != null) return gfReal.Top.Value;
                else return 0.0;
            }
        }

        /// <summary>
        /// Specifies position of the first color at the bottom edge, ranging 0.0 to 1.0. A 0.0 means the position is on the top edge of the cell, and 1.0 means it's on the bottom edge.
        /// This is read-only. Use one of the methods to set the position.
        /// </summary>
        public double GradientBottom
        {
            get
            {
                if (gfReal.Bottom != null) return gfReal.Bottom.Value;
                else return 0.0;
            }
        }

        /// <summary>
        /// Initializes an instance of SLFill. It is recommended to use CreateFill() of the SLDocument class.
        /// </summary>
        public SLFill()
        {
            this.Initialize(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
        }

        internal SLFill(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            this.Initialize(ThemeColors, IndexedColors);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
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
            HasBeenAssignedValues = false;

            RemovePatternFill();
            RemoveGradientFill();

            UsePatternFill = true;
            pfReal.vPatternType = PatternValues.None;
        }

        /// <summary>
        /// Set the foreground color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        public void SetPatternForegroundColor(System.Drawing.Color Color)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.ForegroundColor = Color;
        }

        /// <summary>
        /// Set the foreground color with one of the theme colors.
        /// </summary>
        /// <param name="ColorTheme">The theme color to be used.</param>
        public void SetPatternForegroundColor(SLThemeColorIndexValues ColorTheme)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.SetForegroundThemeColor(ColorTheme);
        }

        /// <summary>
        /// Set the foreground color with one of the theme colors, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="ColorTheme">The theme color to be used.</param>
        /// <param name="ColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPatternForegroundColor(SLThemeColorIndexValues ColorTheme, double ColorTint)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.SetForegroundThemeColor(ColorTheme, ColorTint);
        }

        /// <summary>
        /// Set the background color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        public void SetPatternBackgroundColor(System.Drawing.Color Color)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.BackgroundColor = Color;
        }

        /// <summary>
        /// Set the background color with a theme color.
        /// </summary>
        /// <param name="ColorTheme">The theme color to be used.</param>
        public void SetPatternBackgroundColor(SLThemeColorIndexValues ColorTheme)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.SetBackgroundThemeColor(ColorTheme);
        }

        /// <summary>
        /// Set the background color with one of the theme colors, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="ColorTheme">The theme color to be used.</param>
        /// <param name="ColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPatternBackgroundColor(SLThemeColorIndexValues ColorTheme, double ColorTint)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.SetBackgroundThemeColor(ColorTheme, ColorTint);
        }

        /// <summary>
        /// Set the pattern type.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        public void SetPatternType(PatternValues PatternType)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPattern(PatternValues PatternType, System.Drawing.Color ForegroundColor, System.Drawing.Color BackgroundColor)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
            pfReal.ForegroundColor = ForegroundColor;
            pfReal.BackgroundColor = BackgroundColor;
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPattern(PatternValues PatternType, System.Drawing.Color ForegroundColor, SLThemeColorIndexValues BackgroundColorTheme)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
            pfReal.ForegroundColor = ForegroundColor;
            pfReal.SetBackgroundThemeColor(BackgroundColorTheme);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPattern(PatternValues PatternType, System.Drawing.Color ForegroundColor, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
            pfReal.ForegroundColor = ForegroundColor;
            pfReal.SetBackgroundThemeColor(BackgroundColorTheme, BackgroundColorTint);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPattern(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, System.Drawing.Color BackgroundColor)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
            pfReal.SetForegroundThemeColor(ForegroundColorTheme);
            pfReal.BackgroundColor = BackgroundColor;
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPattern(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, SLThemeColorIndexValues BackgroundColorTheme)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
            pfReal.SetForegroundThemeColor(ForegroundColorTheme);
            pfReal.SetBackgroundThemeColor(BackgroundColorTheme);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPattern(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
            pfReal.SetForegroundThemeColor(ForegroundColorTheme);
            pfReal.SetBackgroundThemeColor(BackgroundColorTheme, BackgroundColorTint);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPattern(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, double ForegroundColorTint, System.Drawing.Color BackgroundColor)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
            pfReal.SetForegroundThemeColor(ForegroundColorTheme, ForegroundColorTint);
            pfReal.BackgroundColor = BackgroundColor;
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPattern(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, double ForegroundColorTint, SLThemeColorIndexValues BackgroundColorTheme)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
            pfReal.SetForegroundThemeColor(ForegroundColorTheme, ForegroundColorTint);
            pfReal.SetBackgroundThemeColor(BackgroundColorTheme);
        }

        /// <summary>
        /// Set the pattern type, foreground color and background color of the fill pattern.
        /// </summary>
        /// <param name="PatternType">The pattern type. Default value is None.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetPattern(PatternValues PatternType, SLThemeColorIndexValues ForegroundColorTheme, double ForegroundColorTint, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = true;
            pfReal.PatternType = PatternType;
            pfReal.SetForegroundThemeColor(ForegroundColorTheme, ForegroundColorTint);
            pfReal.SetBackgroundThemeColor(BackgroundColorTheme, BackgroundColorTint);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1">The first color.</param>
        /// <param name="Color2">The second color.</param>
        public void SetGradient(SLGradientShadingStyleValues ShadingStyle, System.Drawing.Color Color1, System.Drawing.Color Color2)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetGradientFill(ShadingStyle, Color1, Color2);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1">The first color.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        public void SetGradient(SLGradientShadingStyleValues ShadingStyle, System.Drawing.Color Color1, SLThemeColorIndexValues Color2Theme)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetGradientFill(ShadingStyle, Color1, Color2Theme);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1">The first color.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        /// <param name="Color2Tint">The tint applied to the second theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetGradient(SLGradientShadingStyleValues ShadingStyle, System.Drawing.Color Color1, SLThemeColorIndexValues Color2Theme, double Color2Tint)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetGradientFill(ShadingStyle, Color1, Color2Theme, Color2Tint);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color2">The second color.</param>
        public void SetGradient(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, System.Drawing.Color Color2)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetGradientFill(ShadingStyle, Color1Theme, Color2);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        public void SetGradient(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, SLThemeColorIndexValues Color2Theme)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetGradientFill(ShadingStyle, Color1Theme, Color2Theme);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        /// <param name="Color2Tint">The tint applied to the second theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetGradient(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, SLThemeColorIndexValues Color2Theme, double Color2Tint)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetGradientFill(ShadingStyle, Color1Theme, Color2Theme, Color2Tint);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color1Tint">The tint applied to the first theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Color2">The second color.</param>
        public void SetGradient(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, double Color1Tint, System.Drawing.Color Color2)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetGradientFill(ShadingStyle, Color1Theme, Color1Tint, Color2);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color1Tint">The tint applied to the first theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        public void SetGradient(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, double Color1Tint, SLThemeColorIndexValues Color2Theme)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetGradientFill(ShadingStyle, Color1Theme, Color1Tint, Color2Theme);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color1Tint">The tint applied to the first theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        /// <param name="Color2Tint">The tint applied to the second theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetGradient(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, double Color1Tint, SLThemeColorIndexValues Color2Theme, double Color2Tint)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetGradientFill(ShadingStyle, Color1Theme, Color1Tint, Color2Theme, Color2Tint);
        }

        /// <summary>
        /// Set a custom gradient fill. Used in conjunction with AppendGradientStop().
        /// </summary>
        /// <param name="GradientType">The gradient fill type. Default value is Linear.</param>
        /// <param name="Degree">The angle in the direction from which the first color starts. The end color is at 180 degrees from this angle. 0 degrees means start from left, 90 degrees from the top, 180 degrees from the right and 270 degrees from below.</param>
        /// <param name="Left">Specifies position of the first color at the left edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the left edge of the cell, and 1.0 means it's on the right edge.</param>
        /// <param name="Right">Specifies position of the first color at the right edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the left edge of the cell, and 1.0 means it's on the right edge.</param>
        /// <param name="Top">Specifies position of the first color at the top edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the top edge of the cell, and 1.0 means it's on the bottom edge.</param>
        /// <param name="Bottom">Specifies position of the first color at the bottom edge, ranging 0.0 to 1.0. A 0.0 means the position is on the top edge of the cell, and 1.0 means it's on the bottom edge.</param>
        public void SetCustomGradient(GradientValues GradientType, double? Degree, double? Left, double? Right, double? Top, double? Bottom)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.SetCustomGradient(GradientType, Degree, Left, Right, Top, Bottom);
        }

        /// <summary>
        /// Set a gradient stop given a position and a color. Used in conjunction with SetCustomGradient().
        /// </summary>
        /// <param name="Position">Specifies position of the color, ranging from 0.0 to 1.0.</param>
        /// <param name="Color">The color to be used.</param>
        public void AppendGradientStop(double Position, System.Drawing.Color Color)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.AppendGradientStop(Position, Color);
        }

        /// <summary>
        /// Set a gradient stop given a position and a color. Used in conjunction with SetCustomGradient().
        /// </summary>
        /// <param name="Position">Specifies position of the color, ranging from 0.0 to 1.0.</param>
        /// <param name="ColorTheme">The theme color to be used.</param>
        public void AppendGradientStop(double Position, SLThemeColorIndexValues ColorTheme)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.AppendGradientStop(Position, ColorTheme);
        }

        /// <summary>
        /// Set a gradient stop given a position and a color. Used in conjunction with SetCustomGradient().
        /// </summary>
        /// <param name="Position">Specifies position of the color, ranging from 0.0 to 1.0.</param>
        /// <param name="ColorTheme">The theme color to be used.</param>
        /// <param name="ColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void AppendGradientStop(double Position, SLThemeColorIndexValues ColorTheme, double ColorTint)
        {
            HasBeenAssignedValues = true;
            UsePatternFill = false;
            this.gfReal.AppendGradientStop(Position, ColorTheme, ColorTint);
        }

        /// <summary>
        /// Clear all existing gradient stops.
        /// </summary>
        public void ClearGradientStops()
        {
            this.gfReal.ClearGradientStops();
        }

        internal void RemovePatternFill()
        {
            this.pfReal = new SLPatternFill(this.listThemeColors, this.listIndexedColors);
        }

        internal void RemoveGradientFill()
        {
            this.gfReal = new SLGradientFill(this.listThemeColors, this.listIndexedColors);
        }

        internal void FromFill(Fill f)
        {
            this.SetAllNull();

            bool bFound = false;
            if (f.PatternFill != null)
            {
                this.pfReal = new SLPatternFill(this.listThemeColors, this.listIndexedColors);
                this.pfReal.FromPatternFill(f.PatternFill);
                this.UsePatternFill = true;
                bFound = pfReal.HasForegroundColor || pfReal.HasBackgroundColor || pfReal.HasPatternType;
            }
            else if (f.GradientFill != null)
            {
                this.gfReal = new SLGradientFill(this.listThemeColors, this.listIndexedColors);
                this.gfReal.FromGradientFill(f.GradientFill);
                this.UsePatternFill = false;
                bFound = (gfReal.listGradientStops.Count > 0) || gfReal.HasType || gfReal.Degree != null || gfReal.Left != null || gfReal.Right != null || gfReal.Top != null || gfReal.Bottom != null;
            }

            if (bFound)
            {
                HasBeenAssignedValues = true;
            }
            else
            {
                HasBeenAssignedValues = false;

                // must have either PatternFill or GradientFill
                // Default will be an empty PatternFill
                this.pfReal = new SLPatternFill(this.listThemeColors, this.listIndexedColors);
                this.pfReal.PatternType = PatternValues.None;
                this.UsePatternFill = true;

                RemoveGradientFill();
            }
        }

        internal Fill ToFill()
        {
            Fill f = new Fill();
            if (UsePatternFill)
            {
                f.PatternFill = this.pfReal.ToPatternFill();
            }
            else
            {
                f.GradientFill = this.gfReal.ToGradientFill();
            }

            return f;
        }

        internal void FromHash(string Hash)
        {
            Fill f = new Fill();
            f.InnerXml = Hash;
            this.FromFill(f);
        }

        internal string ToHash()
        {
            Fill f = this.ToFill();
            return SLTool.RemoveNamespaceDeclaration(f.InnerXml);
        }

        internal SLFill Clone()
        {
            SLFill f = new SLFill(this.listThemeColors, this.listIndexedColors);
            f.HasBeenAssignedValues = this.HasBeenAssignedValues;
            f.UsePatternFill = this.UsePatternFill;
            f.pfReal = this.pfReal.Clone();
            f.gfReal = this.gfReal.Clone();

            return f;
        }
    }

    /// <summary>
    /// Encapsulates properties and methods for a pattern fill. This simulates the DocumentFormat.OpenXml.Spreadsheet.PatternFill class.
    /// </summary>
    public class SLPatternFill
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        internal bool HasForegroundColor;
        private SLColor clrForegroundColor;
        /// <summary>
        /// The foreground color.
        /// </summary>
        public System.Drawing.Color ForegroundColor
        {
            get { return clrForegroundColor.Color; }
            set
            {
                clrForegroundColor.Color = value;
                HasForegroundColor = (clrForegroundColor.IsEmpty()) ? false : true;
            }
        }

        internal bool HasBackgroundColor;
        private SLColor clrBackgroundColor;
        /// <summary>
        /// The background color.
        /// </summary>
        public System.Drawing.Color BackgroundColor
        {
            get { return clrBackgroundColor.Color; }
            set
            {
                clrBackgroundColor.Color = value;
                HasBackgroundColor = (clrBackgroundColor.Color.IsEmpty) ? false : true;
            }
        }

        internal bool HasPatternType;
        internal PatternValues vPatternType;
        /// <summary>
        /// The pattern type. Default value is None.
        /// </summary>
        public PatternValues PatternType
        {
            get { return vPatternType; }
            set
            {
                vPatternType = value;
                // don't care about the default. If it's set, just use it.
                //HasPatternType = vPatternType != PatternValues.None ? true : false;
                HasPatternType = true;
            }
        }

        /// <summary>
        /// Initializes an instance of SLPatternFill. It is recommended to use CreatePatternFill() of the SLDocument class.
        /// </summary>
        public SLPatternFill()
        {
            this.Initialize(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
        }

        internal SLPatternFill(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            this.Initialize(ThemeColors, IndexedColors);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
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
            this.clrForegroundColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            HasForegroundColor = false;
            this.clrBackgroundColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            HasBackgroundColor = false;
            this.vPatternType = PatternValues.None;
            HasPatternType = false;
        }

        /// <summary>
        /// Set the foreground color with a theme color.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        public void SetForegroundThemeColor(SLThemeColorIndexValues ThemeColorIndex)
        {
            this.clrForegroundColor.SetThemeColor(ThemeColorIndex);
            HasForegroundColor = (clrForegroundColor.Color.IsEmpty) ? false : true;
        }

        /// <summary>
        /// Set the foreground color with a theme color.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetForegroundThemeColor(SLThemeColorIndexValues ThemeColorIndex, double Tint)
        {
            this.clrForegroundColor.SetThemeColor(ThemeColorIndex, Tint);
            HasForegroundColor = (clrForegroundColor.Color.IsEmpty) ? false : true;
        }

        /// <summary>
        /// Set the background color with a theme color.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        public void SetBackgroundThemeColor(SLThemeColorIndexValues ThemeColorIndex)
        {
            this.clrBackgroundColor.SetThemeColor(ThemeColorIndex);
            HasBackgroundColor = (clrBackgroundColor.Color.IsEmpty) ? false : true;
        }

        /// <summary>
        /// Set the background color with a theme color.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetBackgroundThemeColor(SLThemeColorIndexValues ThemeColorIndex, double Tint)
        {
            this.clrBackgroundColor.SetThemeColor(ThemeColorIndex, Tint);
            HasBackgroundColor = (clrBackgroundColor.Color.IsEmpty) ? false : true;
        }

        /// <summary>
        /// Form SLPatternFill from DocumentFormat.OpenXml.Spreadsheet.PatternFill class.
        /// </summary>
        /// <param name="patternFill">The source PatternFill class.</param>
        public void FromPatternFill(PatternFill patternFill)
        {
            this.SetAllNull();
            if (patternFill.ForegroundColor != null)
            {
                this.clrForegroundColor.FromForegroundColor(patternFill.ForegroundColor);
                this.HasForegroundColor = !this.clrForegroundColor.IsEmpty();
            }
            if (patternFill.BackgroundColor != null)
            {
                this.clrBackgroundColor.FromBackgroundColor(patternFill.BackgroundColor);
                this.HasBackgroundColor = !this.clrBackgroundColor.IsEmpty();
            }
            if (patternFill.PatternType != null)
            {
                this.PatternType = patternFill.PatternType;
            }
        }

        /// <summary>
        /// Form a DocumentFormat.OpenXml.Spreadsheet.PatternFill class from SLPatternFill.
        /// </summary>
        /// <returns>A DocumentFormat.OpenXml.Spreadsheet.PatternFill class with the properties of this SLPatternFill class.</returns>
        public PatternFill ToPatternFill()
        {
            PatternFill pf = new PatternFill();
            if (HasForegroundColor) pf.ForegroundColor = this.clrForegroundColor.ToForegroundColor();
            if (HasBackgroundColor) pf.BackgroundColor = this.clrBackgroundColor.ToBackgroundColor();
            if (HasPatternType) pf.PatternType = this.PatternType;

            return pf;
        }

        internal void FromHash(string Hash)
        {
            PatternFill pf = new PatternFill();

            string[] saElementAttribute = Hash.Split(new string[] { SLConstants.XmlPatternFillElementAttributeSeparator }, StringSplitOptions.None);
            if (saElementAttribute.Length >= 2)
            {
                pf.InnerXml = saElementAttribute[0];
                string[] sa = saElementAttribute[1].Split(new string[] { SLConstants.XmlPatternFillAttributeSeparator }, StringSplitOptions.None);
                if (sa.Length >= 1)
                {
                    if (!sa[0].Equals("null")) pf.PatternType = (PatternValues)Enum.Parse(typeof(PatternValues), sa[0]);
                }
            }

            this.FromPatternFill(pf);
        }

        internal string ToHash()
        {
            PatternFill pf = this.ToPatternFill();
            string sXml = SLTool.RemoveNamespaceDeclaration(pf.InnerXml);

            StringBuilder sb = new StringBuilder();

            sb.AppendFormat("{0}{1}", sXml, SLConstants.XmlPatternFillElementAttributeSeparator);

            if (pf.PatternType != null) sb.AppendFormat("{0}{1}", pf.PatternType.Value, SLConstants.XmlPatternFillAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlPatternFillAttributeSeparator);

            return sb.ToString();
        }

        internal SLPatternFill Clone()
        {
            SLPatternFill pf = new SLPatternFill(this.listThemeColors, this.listIndexedColors);
            pf.HasForegroundColor = this.HasForegroundColor;
            pf.clrForegroundColor = this.clrForegroundColor.Clone();
            pf.HasBackgroundColor = this.HasBackgroundColor;
            pf.clrBackgroundColor = this.clrBackgroundColor.Clone();
            pf.HasPatternType = this.HasPatternType;
            pf.vPatternType = this.vPatternType;

            return pf;
        }
    }

    /// <summary>
    /// Encapsulates properties and methods for gradient fills. This simulates the DocumentFormat.OpenXml.Spreadsheet.GradientFill class.
    /// </summary>
    public class SLGradientFill
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        internal List<GradientStop> listGradientStops;

        internal bool HasType;
        private GradientValues vType;
        /// <summary>
        /// The gradient type. Default value is Linear.
        /// </summary>
        public GradientValues Type
        {
            get { return vType; }
            set
            {
                vType = value;
                HasType = vType != GradientValues.Linear ? true : false;
            }
        }

        /// <summary>
        /// The angle in the direction from which the first color starts. The end color is at 180 degrees from this angle. 0 degrees means start from left, 90 degrees from the top, 180 degrees from the right and 270 degrees from below.
        /// </summary>
        public double? Degree { get; set; }

        /// <summary>
        /// Specifies position of the first color at the left edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the left edge of the cell, and 1.0 means it's on the right edge.
        /// </summary>
        public double? Left { get; set; }

        /// <summary>
        /// Specifies position of the first color at the right edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the left edge of the cell, and 1.0 means it's on the right edge.
        /// </summary>
        public double? Right { get; set; }

        /// <summary>
        /// Specifies position of the first color at the top edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the top edge of the cell, and 1.0 means it's on the bottom edge.
        /// </summary>
        public double? Top { get; set; }

        /// <summary>
        /// Specifies position of the first color at the bottom edge, ranging 0.0 to 1.0. A 0.0 means the position is on the top edge of the cell, and 1.0 means it's on the bottom edge.
        /// </summary>
        public double? Bottom { get; set; }

        /// <summary>
        /// Initializes an instance of SLGradientFill. It is recommended to use CreateGradientFill() of the SLDocument class.
        /// </summary>
        public SLGradientFill()
        {
            this.Initialize(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
        }

        internal SLGradientFill(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            this.Initialize(ThemeColors, IndexedColors);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
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
            this.listGradientStops = new List<GradientStop>();
            this.vType = GradientValues.Linear;
            this.HasType = false;
            this.Degree = null;
            this.Left = null;
            this.Right = null;
            this.Top = null;
            this.Bottom = null;
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1">The first color.</param>
        /// <param name="Color2">The second color.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, System.Drawing.Color Color1, System.Drawing.Color Color2)
        {
            SLColor clr1 = new SLColor(this.listThemeColors, this.listIndexedColors);
            SLColor clr2 = new SLColor(this.listThemeColors, this.listIndexedColors);

            clr1.Rgb = string.Format("{0}{1}{2}", Color1.R.ToString("x2"), Color1.G.ToString("x2"), Color1.B.ToString("x2"));
            clr2.Rgb = string.Format("{0}{1}{2}", Color2.R.ToString("x2"), Color2.G.ToString("x2"), Color2.B.ToString("x2"));

            SetGradientFill(ShadingStyle, clr1, clr2);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1">The first color.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, System.Drawing.Color Color1, SLThemeColorIndexValues Color2Theme)
        {
            SLColor clr1 = new SLColor(this.listThemeColors, this.listIndexedColors);
            SLColor clr2 = new SLColor(this.listThemeColors, this.listIndexedColors);

            clr1.Rgb = string.Format("{0}{1}{2}", Color1.R.ToString("x2"), Color1.G.ToString("x2"), Color1.B.ToString("x2"));
            clr2.SetThemeColor(Color2Theme);

            SetGradientFill(ShadingStyle, clr1, clr2);
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
            SLColor clr1 = new SLColor(this.listThemeColors, this.listIndexedColors);
            SLColor clr2 = new SLColor(this.listThemeColors, this.listIndexedColors);

            clr1.Rgb = string.Format("{0}{1}{2}", Color1.R.ToString("x2"), Color1.G.ToString("x2"), Color1.B.ToString("x2"));
            clr2.SetThemeColor(Color2Theme, Color2Tint);

            SetGradientFill(ShadingStyle, clr1, clr2);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color2">The second color.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, System.Drawing.Color Color2)
        {
            SLColor clr1 = new SLColor(this.listThemeColors, this.listIndexedColors);
            SLColor clr2 = new SLColor(this.listThemeColors, this.listIndexedColors);

            clr1.SetThemeColor(Color1Theme);
            clr2.Rgb = string.Format("{0}{1}{2}", Color2.R.ToString("x2"), Color2.G.ToString("x2"), Color2.B.ToString("x2"));

            SetGradientFill(ShadingStyle, clr1, clr2);
        }

        /// <summary>
        /// Set a gradient fill given the shading style and 2 colors.
        /// </summary>
        /// <param name="ShadingStyle">The gradient shading style.</param>
        /// <param name="Color1Theme">The first color as one of the theme colors.</param>
        /// <param name="Color2Theme">The second color as one of the theme colors.</param>
        public void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, SLThemeColorIndexValues Color1Theme, SLThemeColorIndexValues Color2Theme)
        {
            SLColor clr1 = new SLColor(this.listThemeColors, this.listIndexedColors);
            SLColor clr2 = new SLColor(this.listThemeColors, this.listIndexedColors);

            clr1.SetThemeColor(Color1Theme);
            clr2.SetThemeColor(Color2Theme);

            SetGradientFill(ShadingStyle, clr1, clr2);
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
            SLColor clr1 = new SLColor(this.listThemeColors, this.listIndexedColors);
            SLColor clr2 = new SLColor(this.listThemeColors, this.listIndexedColors);

            clr1.SetThemeColor(Color1Theme);
            clr2.SetThemeColor(Color2Theme, Color2Tint);

            SetGradientFill(ShadingStyle, clr1, clr2);
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
            SLColor clr1 = new SLColor(this.listThemeColors, this.listIndexedColors);
            SLColor clr2 = new SLColor(this.listThemeColors, this.listIndexedColors);

            clr1.SetThemeColor(Color1Theme, Color1Tint);
            clr2.Rgb = string.Format("{0}{1}{2}", Color2.R.ToString("x2"), Color2.G.ToString("x2"), Color2.B.ToString("x2"));

            SetGradientFill(ShadingStyle, clr1, clr2);
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
            SLColor clr1 = new SLColor(this.listThemeColors, this.listIndexedColors);
            SLColor clr2 = new SLColor(this.listThemeColors, this.listIndexedColors);

            clr1.SetThemeColor(Color1Theme, Color1Tint);
            clr2.SetThemeColor(Color2Theme);

            SetGradientFill(ShadingStyle, clr1, clr2);
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
            SLColor clr1 = new SLColor(this.listThemeColors, this.listIndexedColors);
            SLColor clr2 = new SLColor(this.listThemeColors, this.listIndexedColors);

            clr1.SetThemeColor(Color1Theme, Color1Tint);
            clr2.SetThemeColor(Color2Theme, Color2Tint);

            SetGradientFill(ShadingStyle, clr1, clr2);
        }

        private void SetGradientFill(SLGradientShadingStyleValues ShadingStyle, SLColor Color1, SLColor Color2)
        {
            GradientStop gs;

            switch (ShadingStyle)
            {
                case SLGradientShadingStyleValues.Horizontal1:
                    this.Degree = 90;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.Horizontal2:
                    this.Degree = 270;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.Horizontal3:
                    this.Degree = 90;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 0.5;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.Vertical1:
                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.Vertical2:
                    this.Degree = 180;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.Vertical3:
                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 0.5;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.DiagonalUp1:
                    this.Degree = 45;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.DiagonalUp2:
                    this.Degree = 225;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.DiagonalUp3:
                    this.Degree = 45;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 0.5;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.DiagonalDown1:
                    this.Degree = 135;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.DiagonalDown2:
                    this.Degree = 315;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.DiagonalDown3:
                    this.Degree = 135;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 0.5;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.Corner1:
                    this.Type = GradientValues.Path;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.Corner2:
                    this.Type = GradientValues.Path;
                    this.Left = 1;
                    this.Right = 1;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.Corner3:
                    this.Type = GradientValues.Path;
                    this.Top = 1;
                    this.Bottom = 1;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.Corner4:
                    this.Type = GradientValues.Path;
                    this.Left = 1;
                    this.Right = 1;
                    this.Top = 1;
                    this.Bottom = 1;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
                case SLGradientShadingStyleValues.FromCenter:
                    this.Type = GradientValues.Path;
                    this.Left = 0.5;
                    this.Right = 0.5;
                    this.Top = 0.5;
                    this.Bottom = 0.5;

                    gs = new GradientStop();
                    gs.Position = 0;
                    gs.Color = Color1.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);

                    gs = new GradientStop();
                    gs.Position = 1;
                    gs.Color = Color2.ToSpreadsheetColor();
                    this.listGradientStops.Add(gs);
                    break;
            }
        }

        /// <summary>
        /// Clear all existing gradient stops.
        /// </summary>
        public void ClearGradientStops()
        {
            this.listGradientStops.Clear();
        }

        /// <summary>
        /// Set a custom gradient fill. Used in conjunction with AppendGradientStop().
        /// </summary>
        /// <param name="GradientType">The gradient fill type. Default value is Linear.</param>
        /// <param name="Degree">The angle in the direction from which the first color starts. The end color is at 180 degrees from this angle. 0 degrees means start from left, 90 degrees from the top, 180 degrees from the right and 270 degrees from below.</param>
        /// <param name="Left">Specifies position of the first color at the left edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the left edge of the cell, and 1.0 means it's on the right edge.</param>
        /// <param name="Right">Specifies position of the first color at the right edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the left edge of the cell, and 1.0 means it's on the right edge.</param>
        /// <param name="Top">Specifies position of the first color at the top edge, ranging from 0.0 to 1.0. A 0.0 means the position is on the top edge of the cell, and 1.0 means it's on the bottom edge.</param>
        /// <param name="Bottom">Specifies position of the first color at the bottom edge, ranging 0.0 to 1.0. A 0.0 means the position is on the top edge of the cell, and 1.0 means it's on the bottom edge.</param>
        public void SetCustomGradient(GradientValues GradientType, double? Degree, double? Left, double? Right, double? Top, double? Bottom)
        {
            this.Type = GradientType;
            if (Degree != null) this.Degree = Degree.Value;
            if (Left != null) this.Left = Left.Value;
            if (Right != null) this.Right = Right.Value;
            if (Top != null) this.Top = Top.Value;
            if (Bottom != null) this.Bottom = Bottom.Value;
        }

        /// <summary>
        /// Set a gradient stop given a position and a color. Used in conjunction with SetCustomGradient().
        /// </summary>
        /// <param name="Position">Specifies position of the color, ranging from 0.0 to 1.0.</param>
        /// <param name="Color">The color to be used.</param>
        public void AppendGradientStop(double Position, System.Drawing.Color Color)
        {
            SLColor clr = new SLColor(this.listThemeColors, this.listIndexedColors);
            clr.Rgb = string.Format("{0}{1}{2}", Color.R.ToString("x2"), Color.G.ToString("x2"), Color.B.ToString("x2"));
            GradientStop gs = new GradientStop();
            gs.Position = Position;
            gs.Color = clr.ToSpreadsheetColor();
            listGradientStops.Add(gs);
        }

        /// <summary>
        /// Set a gradient stop given a position and a color. Used in conjunction with SetCustomGradient().
        /// </summary>
        /// <param name="Position">Specifies position of the color, ranging from 0.0 to 1.0.</param>
        /// <param name="ColorTheme">The theme color to be used.</param>
        public void AppendGradientStop(double Position, SLThemeColorIndexValues ColorTheme)
        {
            SLColor clr = new SLColor(this.listThemeColors, this.listIndexedColors);
            clr.SetThemeColor(ColorTheme);
            GradientStop gs = new GradientStop();
            gs.Position = Position;
            gs.Color = clr.ToSpreadsheetColor();
            listGradientStops.Add(gs);
        }

        /// <summary>
        /// Set a gradient stop given a position and a color. Used in conjunction with SetCustomGradient().
        /// </summary>
        /// <param name="Position">Specifies position of the color, ranging from 0.0 to 1.0.</param>
        /// <param name="ColorTheme">The theme color to be used.</param>
        /// <param name="ColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void AppendGradientStop(double Position, SLThemeColorIndexValues ColorTheme, double ColorTint)
        {
            SLColor clr = new SLColor(this.listThemeColors, this.listIndexedColors);
            clr.SetThemeColor(ColorTheme, ColorTint);
            GradientStop gs = new GradientStop();
            gs.Position = Position;
            gs.Color = clr.ToSpreadsheetColor();
            listGradientStops.Add(gs);
        }

        /// <summary>
        /// Form SLGradientFill from DocumentFormat.OpenXml.Spreadsheet.GradientFill class.
        /// </summary>
        /// <param name="gradientFill">The source DocumentFormat.OpenXml.Spreadsheet.GradientFill class.</param>
        public void FromGradientFill(GradientFill gradientFill)
        {
            this.SetAllNull();

            using (OpenXmlReader oxr = OpenXmlReader.Create(gradientFill))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(GradientStop))
                    {
                        listGradientStops.Add((GradientStop)oxr.LoadCurrentElement().CloneNode(true));
                    }
                }
            }

            if (gradientFill.Type != null)
            {
                this.Type = gradientFill.Type.Value;
            }

            if (gradientFill.Degree != null)
            {
                this.Degree = gradientFill.Degree.Value;
            }

            if (gradientFill.Left != null)
            {
                this.Left = gradientFill.Left.Value;
            }

            if (gradientFill.Right != null)
            {
                this.Right = gradientFill.Right.Value;
            }

            if (gradientFill.Top != null)
            {
                this.Top = gradientFill.Top.Value;
            }

            if (gradientFill.Bottom != null)
            {
                this.Bottom = gradientFill.Bottom.Value;
            }
        }

        /// <summary>
        /// Form a DocumentFormat.OpenXml.Spreadsheet.GradientFill class from SLGradientFill.
        /// </summary>
        /// <returns>A DocumentFormat.OpenXml.Spreadsheet.GradientFill with the properties of this SLGradientFill.</returns>
        public GradientFill ToGradientFill()
        {
            GradientFill gf = new GradientFill();
            for (int i = 0; i < this.listGradientStops.Count; ++i)
            {
                gf.Append(listGradientStops[i]);
            }

            if (HasType) gf.Type = this.Type;
            if (this.Degree != null) gf.Degree = this.Degree.Value;
            if (this.Left != null) gf.Left = this.Left.Value;
            if (this.Right != null) gf.Right = this.Right.Value;
            if (this.Top != null) gf.Top = this.Top.Value;
            if (this.Bottom != null) gf.Bottom = this.Bottom.Value;

            return gf;
        }

        internal void FromHash(string Hash)
        {
            GradientFill gf = new GradientFill();

            string[] saElementAttribute = Hash.Split(new string[] { SLConstants.XmlGradientFillElementAttributeSeparator }, StringSplitOptions.None);
            if (saElementAttribute.Length >= 2)
            {
                gf.InnerXml = saElementAttribute[0];
                string[] sa = saElementAttribute[1].Split(new string[] { SLConstants.XmlGradientFillAttributeSeparator }, StringSplitOptions.None);
                if (sa.Length >= 6)
                {
                    if (!sa[0].Equals("null")) gf.Type = (GradientValues)Enum.Parse(typeof(GradientValues), sa[0]);

                    if (!sa[1].Equals("null")) gf.Degree = double.Parse(sa[1]);

                    if (!sa[2].Equals("null")) gf.Left = double.Parse(sa[2]);

                    if (!sa[3].Equals("null")) gf.Right = double.Parse(sa[3]);

                    if (!sa[4].Equals("null")) gf.Top = double.Parse(sa[4]);

                    if (!sa[5].Equals("null")) gf.Bottom = double.Parse(sa[5]);
                }
            }

            this.FromGradientFill(gf);
        }

        internal string ToHash()
        {
            GradientFill gf = this.ToGradientFill();
            string sXml = SLTool.RemoveNamespaceDeclaration(gf.InnerXml);

            StringBuilder sb = new StringBuilder();

            sb.AppendFormat("{0}{1}", sXml, SLConstants.XmlGradientFillElementAttributeSeparator);

            if (gf.Type != null) sb.AppendFormat("{0}{1}", gf.Type.Value, SLConstants.XmlGradientFillAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlGradientFillAttributeSeparator);

            if (gf.Degree != null) sb.AppendFormat("{0}{1}", gf.Degree.Value, SLConstants.XmlGradientFillAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlGradientFillAttributeSeparator);

            if (gf.Left != null) sb.AppendFormat("{0}{1}", gf.Left.Value, SLConstants.XmlGradientFillAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlGradientFillAttributeSeparator);

            if (gf.Right != null) sb.AppendFormat("{0}{1}", gf.Right.Value, SLConstants.XmlGradientFillAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlGradientFillAttributeSeparator);

            if (gf.Top != null) sb.AppendFormat("{0}{1}", gf.Top.Value, SLConstants.XmlGradientFillAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlGradientFillAttributeSeparator);

            if (gf.Bottom != null) sb.AppendFormat("{0}{1}", gf.Bottom.Value, SLConstants.XmlGradientFillAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlGradientFillAttributeSeparator);

            return sb.ToString();
        }

        internal SLGradientFill Clone()
        {
            SLGradientFill gf = new SLGradientFill(this.listThemeColors, this.listIndexedColors);

            gf.listGradientStops = new List<GradientStop>();
            for (int i = 0; i < this.listGradientStops.Count; ++i)
            {
                gf.listGradientStops.Add((GradientStop)this.listGradientStops[i].CloneNode(true));
            }

            gf.HasType = this.HasType;
            gf.vType = this.vType;
            gf.Degree = this.Degree;
            gf.Left = this.Left;
            gf.Right = this.Right;
            gf.Top = this.Top;
            gf.Bottom = this.Bottom;

            return gf;
        }
    }
}
