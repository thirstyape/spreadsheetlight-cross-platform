using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

// Targetting .NET Standard means Forms is no longer useful
// Keeping this as a comment as a reminder of pre-Standard days
// (or until this comment is no longer useful :)
//using System.Windows.Forms;

namespace SpreadsheetLight
{
    /// <summary>
    /// Specifies the built-in theme type.
    /// </summary>
    public enum SLThemeTypeValues
    {
        /// <summary>
        /// Office theme with Cambria and Calibri as the major and minor fonts respectively.
        /// </summary>
        Office = 0,
        /// <summary>
        /// Office2013 theme with Calibri Light and Calibri as the major and minor fonts respectively.
        /// </summary>
        Office2013,
        /// <summary>
        /// Adjacency theme with Cambria and Calibri as the major and minor fonts respectively.
        /// </summary>
        Adjacency,
        /// <summary>
        /// Angles theme with Franklin Gothic Medium and Franklin Gothic Book as the major and minor fonts respectively.
        /// </summary>
        Angles,
        /// <summary>
        /// Apex theme with Lucida Sans and Book Antiqua as the major and minor fonts respectively.
        /// </summary>
        Apex,
        /// <summary>
        /// Apothecary theme with Book Antiqua and Century Gothic as the major and minor fonts respectively.
        /// </summary>
        Apothecary,
        /// <summary>
        /// Aspect theme with Verdana as both major and minor fonts.
        /// </summary>
        Aspect,
        /// <summary>
        /// Austin theme with Century Gothic as both major and minor fonts.
        /// </summary>
        Austin,
        /// <summary>
        /// Black Tie theme with Garamond as both major and minor fonts.
        /// </summary>
        BlackTie,
        /// <summary>
        /// Civic theme with Georgia as both major and minor fonts.
        /// </summary>
        Civic,
        /// <summary>
        /// Clarity theme with Arial as both major and minor fonts.
        /// </summary>
        Clarity,
        /// <summary>
        /// Composite theme with Calibri as both major and minor fonts.
        /// </summary>
        Composite,
        /// <summary>
        /// Concourse theme with Lucida Sans Unicode as both major and minor fonts.
        /// </summary>
        Concourse,
        /// <summary>
        /// Couture theme with Garamond as both major and minor fonts.
        /// </summary>
        Couture,
        /// <summary>
        /// Elemental theme with Palatino Linotype as both major and minor fonts.
        /// </summary>
        Elemental,
        /// <summary>
        /// Equity theme with Franklin Gothic Book and Perpetua as the major and minor fonts respectively.
        /// </summary>
        Equity,
        /// <summary>
        /// Essential theme with Arial Black and Arial as the major and minor fonts respectively.
        /// </summary>
        Essential,
        /// <summary>
        /// Executive theme with Century Gothic and Palatino Linotype as the major and minor fonts respectively.
        /// </summary>
        Executive,
        /// <summary>
        /// Facet theme with Trebuchet MS as both major and minor fonts.
        /// </summary>
        Facet,
        /// <summary>
        /// Flow theme with Calibri and Constantia as the major and minor fonts respectively.
        /// </summary>
        Flow,
        /// <summary>
        /// Foundry theme with Rockwell as both major and minor fonts.
        /// </summary>
        Foundry,
        /// <summary>
        /// Grid theme with Franklin Gothic Medium as both major and minor fonts.
        /// </summary>
        Grid,
        /// <summary>
        /// Hardcover theme with Book Antiqua as both major and minor fonts.
        /// </summary>
        Hardcover,
        /// <summary>
        /// Horizon theme with Arial Narrow as both major and minor fonts.
        /// </summary>
        Horizon,
        /// <summary>
        /// Integral theme with Tw Cen MT Condensed and Tw Cen MT as the major and minor fonts respectively.
        /// </summary>
        Integral,
        /// <summary>
        /// Ion theme with Century Gothic as both major and minor fonts.
        /// </summary>
        Ion,
        /// <summary>
        /// Ion Boardroom theme with Century Gothic as both major and minor fonts.
        /// </summary>
        IonBoardroom,
        /// <summary>
        /// Median theme with Tw Cen MT as both major and minor fonts.
        /// </summary>
        Median,
        /// <summary>
        /// Metro theme with Consolas and Corbel as the major and minor fonts respectively.
        /// </summary>
        Metro,
        /// <summary>
        /// Module theme with Corbel as both major and minor fonts.
        /// </summary>
        Module,
        /// <summary>
        /// Newsprint theme with Impact and Times New Roman as the major and minor fonts respectively.
        /// </summary>
        Newsprint,
        /// <summary>
        /// Opulent theme with Trebuchet MS as both major and minor fonts.
        /// </summary>
        Opulent,
        /// <summary>
        /// Organic theme with Garamond as both major and minor fonts.
        /// </summary>
        Organic,
        /// <summary>
        /// Oriel theme with Century Schoolbook as both major and minor fonts.
        /// </summary>
        Oriel,
        /// <summary>
        /// Origin theme with Bookman Old Style and Gill Sans MT as the major and minor fonts respectively.
        /// </summary>
        Origin,
        /// <summary>
        /// Paper theme with Constantia as both major and minor fonts.
        /// </summary>
        Paper,
        /// <summary>
        /// Perspective theme with Arial as both major and minor fonts.
        /// </summary>
        Perspective,
        /// <summary>
        /// Pushpin theme with Constantia and Franklin Gothic Book as the major and minor fonts respectively.
        /// </summary>
        Pushpin,
        /// <summary>
        /// Retrospect theme with Calibri Light and Calibri as the major and minor fonts respectively.
        /// </summary>
        Retrospect,
        /// <summary>
        /// Slice theme with Century Gothic as both major and minor fonts.
        /// </summary>
        Slice,
        /// <summary>
        /// Slipstream theme with Trebuchet MS as both major and minor fonts.
        /// </summary>
        Slipstream,
        /// <summary>
        /// Solstice theme with Gill Sans MT as both major and minor fonts.
        /// </summary>
        Solstice,
        /// <summary>
        /// Technic theme with Franklin Gothic Book and Arial as the major and minor fonts respectively.
        /// </summary>
        Technic,
        /// <summary>
        /// Thatch theme with Tw Cen MT as both major and minor fonts.
        /// </summary>
        Thatch,
        /// <summary>
        /// Trek theme with Franklin Gothic Medium and Franklin Gothic Book as the major and minor fonts respectively.
        /// </summary>
        Trek,
        /// <summary>
        /// Urban theme with Trebuchet MS and Georgia as the major and minor fonts respectively.
        /// </summary>
        Urban,
        /// <summary>
        /// Verve theme with Century Gothic as both major and minor fonts.
        /// </summary>
        Verve,
        /// <summary>
        /// Waveform theme with Candara as both major and minor fonts.
        /// </summary>
        Waveform,
        /// <summary>
        /// Wisp theme with Century Gothic as both major and minor fonts.
        /// </summary>
        Wisp,
        /// <summary>
        /// Autumn theme with Verdana as both major and minor fonts.
        /// </summary>
        Autumn,
        /// <summary>
        /// Banded theme with Corbel as both major and minor fonts.
        /// </summary>
        Banded,
        /// <summary>
        /// Basis theme with Corbel as both major and minor fonts.
        /// </summary>
        Basis,
        /// <summary>
        /// Berlin theme with Trebuchet MS as both major and minor fonts.
        /// </summary>
        Berlin,
        /// <summary>
        /// Celestial theme with Calibri Light and Calibri as the major and minor fonts respectively.
        /// </summary>
        Celestial,
        /// <summary>
        /// Circuit theme with Tw Cen MT as both major and minor fonts.
        /// </summary>
        Circuit,
        /// <summary>
        /// Damask theme Bookman Old Style and Rockwell as the major and minor fonts respectively.
        /// </summary>
        Damask,
        /// <summary>
        /// Decatur theme with Bodoni MT Condensed and Franklin Gothic Book as the major and minor fonts respectively.
        /// </summary>
        Decatur,
        /// <summary>
        /// Depth theme with Corbel as both major and minor fonts.
        /// </summary>
        Depth,
        /// <summary>
        /// Dividend theme with Gill Sans MT as both major and minor fonts.
        /// </summary>
        Dividend,
        /// <summary>
        /// Droplet theme with Tw Cen MT as both major and minor fonts.
        /// </summary>
        Droplet,
        /// <summary>
        /// Frame theme with Corbel as both major and minor fonts.
        /// </summary>
        Frame,
        /// <summary>
        /// Kilter theme with Rockwell as both major and minor fonts.
        /// </summary>
        Kilter,
        /// <summary>
        /// Macro theme with Calibri as both major and minor fonts.
        /// </summary>
        Macro,
        /// <summary>
        /// Main Event theme with Impact as both major and minor fonts.
        /// </summary>
        MainEvent,
        /// <summary>
        /// Mesh theme with Century Gothic as both major and minor fonts.
        /// </summary>
        Mesh,
        /// <summary>
        /// Metropolitan theme with Calibri Light as both major and minor fonts.
        /// </summary>
        Metropolitan,
        /// <summary>
        /// Mylar theme with Corbel as both major and minor fonts.
        /// </summary>
        Mylar,
        /// <summary>
        /// Parallax theme with Corbel as both major and minor fonts.
        /// </summary>
        Parallax,
        /// <summary>
        /// Quotable theme with Century Gothic as both major and minor fonts.
        /// </summary>
        Quotable,
        /// <summary>
        /// Savon theme with Century Gothic as both major and minor fonts.
        /// </summary>
        Savon,
        /// <summary>
        /// Sketchbook theme with Cambria as both major and minor fonts.
        /// </summary>
        Sketchbook,
        /// <summary>
        /// Slate theme with Calisto MT as both major and minor fonts.
        /// </summary>
        Slate,
        /// <summary>
        /// Soho theme with Candara as both major and minor fonts.
        /// </summary>
        Soho,
        /// <summary>
        /// Spring theme with Verdana as both major and minor fonts.
        /// </summary>
        Spring,
        /// <summary>
        /// Summer theme with Verdana as both major and minor fonts.
        /// </summary>
        Summer,
        /// <summary>
        /// Thermal theme with Calibri as both major and minor fonts.
        /// </summary>
        Thermal,
        /// <summary>
        /// Tradeshow theme with Arial Black and Candara as the major and minor fonts respectively.
        /// </summary>
        Tradeshow,
        /// <summary>
        /// Urban Pop theme with Gill Sans MT as both major and minor fonts.
        /// </summary>
        UrbanPop,
        /// <summary>
        /// Vapor Trail theme with Century Gothic as both major and minor fonts.
        /// </summary>
        VaporTrail,
        /// <summary>
        /// View theme with Century Schoolbook as both major and minor fonts.
        /// </summary>
        View,
        /// <summary>
        /// Winter theme with Verdana as both major and minor fonts.
        /// </summary>
        Winter,
        /// <summary>
        /// Wood Type theme with Rockwell Condensed and Rockwell as the major and minor fonts respectively.
        /// </summary>
        WoodType
    }

    // even though it's Dark1, Light1, Dark2, Light2 in the XML
    // the indexing uses Light1, Dark1, Light2, Dark2 (and then the accents)
    // Don't know why Excel and the underlying Open XML theme indexing is inconsistent...
    /// <summary>
    /// Specifies the theme color type.
    /// </summary>
    public enum SLThemeColorIndexValues
    {
        /// <summary>
        /// Typically pure white. For convenience, this also doubles as "Background 1".
        /// </summary>
        Light1Color = 0,
        /// <summary>
        /// Typically pure black. For convenience, this also doubles as "Text 1".
        /// </summary>
        Dark1Color,
        /// <summary>
        /// A light color that still has visual contrast against dark tints of the accent colors. For convenience, this also doubles as "Background 2".
        /// </summary>
        Light2Color,
        /// <summary>
        /// A dark color that still has visual contrast against light tints of the accent colors. For convenience, this also doubles as "Text 2".
        /// </summary>
        Dark2Color,
        /// <summary>
        /// Accent1 color
        /// </summary>
        Accent1Color,
        /// <summary>
        /// Accent2 color
        /// </summary>
        Accent2Color,
        /// <summary>
        /// Accent3 color
        /// </summary>
        Accent3Color,
        /// <summary>
        /// Accent4 color
        /// </summary>
        Accent4Color,
        /// <summary>
        /// Accent5 color
        /// </summary>
        Accent5Color,
        /// <summary>
        /// Accent6 color
        /// </summary>
        Accent6Color,
        /// <summary>
        /// Color of a hyperlink
        /// </summary>
        Hyperlink,
        /// <summary>
        /// Color of a followed hyperlink
        /// </summary>
        FollowedHyperlinkColor
    }

    internal class SLSimpleTheme
    {
        internal List<System.Drawing.Color> listIndexedColors;
        internal List<System.Drawing.Color> listThemeColors;
        internal List<string> listThemeColorsHex;
        internal SLThemeTypeValues InternalThemeType = SLThemeTypeValues.Office;

        internal bool ThrowExceptionsIfAny { get; set; }

        private string sThemeName;
        internal string ThemeName
        {
            get { return sThemeName; }
        }

        private string sMajorLatinFont;
        internal string MajorLatinFont
        {
            get { return sMajorLatinFont; }
        }

        private string sMinorLatinFont;
        internal string MinorLatinFont
        {
            get { return sMinorLatinFont; }
        }

        private double fThemeColumnWidth;
        internal double ThemeColumnWidth
        {
            get { return fThemeColumnWidth; }
        }

        private int iThemeMaxDigitWidth;
        internal int ThemeMaxDigitWidth
        {
            get { return iThemeMaxDigitWidth; }
        }

        private long lThemeColumnWidthInEMU;
        internal long ThemeColumnWidthInEMU
        {
            get { return lThemeColumnWidthInEMU; }
        }

        private double fThemeRowHeight;
        internal double ThemeRowHeight
        {
            get { return fThemeRowHeight; }
        }

        private long lThemeRowHeightInEMU;
        internal long ThemeRowHeightInEMU
        {
            get { return lThemeRowHeightInEMU; }
        }

        internal List<double> listColumnStepSize;

        internal SLSimpleTheme(WorkbookPart wbp, SLThemeTypeValues themetype, bool ThrowExceptionsIfAny)
        {
            this.ThrowExceptionsIfAny = ThrowExceptionsIfAny;

            LoadIndexedColors(wbp);
            InitialiseThemeColors();
            InternalThemeType = themetype;

            bool bHasTheme = (wbp.ThemePart != null) ? true : false;
            if (bHasTheme)
            {
                // load in default values in case the theme file has missing values
                LoadBuiltinTheme(SLThemeTypeValues.Office);
                LoadTheme(wbp);
            }
            else
            {
                LoadBuiltinTheme(themetype);
            }

            CalculateRowColumnInfo();
        }

        internal SLSimpleTheme(WorkbookPart wbp, SLThemeSettings Settings, bool ThrowExceptionsIfAny)
        {
            this.ThrowExceptionsIfAny = ThrowExceptionsIfAny;

            LoadIndexedColors(wbp);
            InitialiseThemeColors();
            InternalThemeType = SLThemeTypeValues.Office;

            bool bHasTheme = (wbp.ThemePart != null) ? true : false;
            if (bHasTheme)
            {
                // load in default values in case the theme file has missing values
                LoadBuiltinTheme(SLThemeTypeValues.Office);
                LoadTheme(wbp);
            }
            else
            {
                LoadBuiltinTheme(SLThemeTypeValues.Office);

                this.sThemeName = Settings.ThemeName;
                this.sMajorLatinFont = Settings.MajorLatinFont;
                this.sMinorLatinFont = Settings.MinorLatinFont;

                listThemeColors[(int)SLThemeColorIndexValues.Dark1Color] = Settings.Dark1Color;
                listThemeColors[(int)SLThemeColorIndexValues.Light1Color] = Settings.Light1Color;
                listThemeColors[(int)SLThemeColorIndexValues.Dark2Color] = Settings.Dark2Color;
                listThemeColors[(int)SLThemeColorIndexValues.Light2Color] = Settings.Light2Color;
                listThemeColors[(int)SLThemeColorIndexValues.Accent1Color] = Settings.Accent1Color;
                listThemeColors[(int)SLThemeColorIndexValues.Accent2Color] = Settings.Accent2Color;
                listThemeColors[(int)SLThemeColorIndexValues.Accent3Color] = Settings.Accent3Color;
                listThemeColors[(int)SLThemeColorIndexValues.Accent4Color] = Settings.Accent4Color;
                listThemeColors[(int)SLThemeColorIndexValues.Accent5Color] = Settings.Accent5Color;
                listThemeColors[(int)SLThemeColorIndexValues.Accent6Color] = Settings.Accent6Color;
                listThemeColors[(int)SLThemeColorIndexValues.Hyperlink] = Settings.Hyperlink;
                listThemeColors[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = Settings.FollowedHyperlinkColor;

                for (int i = 0; i < listThemeColors.Count; ++i)
                {
                    listThemeColorsHex[i] = string.Format("{0}{1}{2}", listThemeColors[i].R.ToString("x2"), listThemeColors[i].G.ToString("x2"), listThemeColors[i].B.ToString("x2"));
                }
            }

            CalculateRowColumnInfo();
        }

        internal void InitialiseThemeColors()
        {
            listThemeColors = new List<System.Drawing.Color>();
            listThemeColorsHex = new List<string>();
            for (int i = 0; i < 12; ++i)
            {
                listThemeColors.Add(System.Drawing.Color.White);
                listThemeColorsHex.Add("FFFFFF");
            }
        }

        internal void CalculateRowColumnInfo()
        {
            System.Drawing.Font usablefont = SLTool.GetUsableNormalFont(this.MinorLatinFont, SLConstants.DefaultFontSize, System.Drawing.FontStyle.Regular, this.ThrowExceptionsIfAny);

            // WARNING: The following algorithm is not guaranteed to work for all fonts.
            // But any algorithm is better than *no* algorithm.
            // This is tested for all 32 typefaces of the 74 built-in themes at both
            // 96 DPI and 120 DPI.
            // Huh, Verdana? It's the exception, with its own exception code.
            // It works great for a web design typeface and looks good on the screen, but
            // man I hate calculating font metrics on Verdana...
            // Rockwell Condensed and Tw Cen MT Condensed aren't as irritating...

            // Office2013 additions: Calibri Light, Rockwell Condensed, Tw Cen MT Condensed

            // What are the 32 typefaces? Alright fine...
            // Arial, Arial Black, Arial Narrow, Bodoni MT Condensed, Book Antiqua, Bookman Old Style,
            // Calibri Light, Calibri, Cambria, Candara, Century Gothic, Century Schoolbook, Consolas, Constantia,
            // Corbel, Franklin Gothic Book, Franklin Gothic Medium, Garamond, Georgia, Gill Sans MT,
            // Impact, Lucida Sans, Lucida Sans Unicode, Palatino Linotype, Perpetua, Rockwell Condensed, Rockwell,
            // Times New Roman, Trebuchet MS, Tw Cen MT Condensed, Tw Cen MT, and *grrr* Verdana

            // Let's have a version with the double quotes because I'm tired of typing when I need
            // to do more empirical experiments...
            // "Arial", "Arial Black", "Arial Narrow", "Bodoni MT Condensed", "Book Antiqua", "Bookman Old Style",
            // "Calibri Light", "Calibri", "Cambria", "Candara", "Century Gothic", "Century Schoolbook", "Consolas", "Constantia",
            // "Corbel", "Franklin Gothic Book", "Franklin Gothic Medium", "Garamond", "Georgia", "Gill Sans MT",
            // "Impact", "Lucida Sans", "Lucida Sans Unicode", "Palatino Linotype", "Perpetua", "Rockwell Condensed", "Rockwell",
            // "Times New Roman", "Trebuchet MS", "Tw Cen MT Condensed", "Tw Cen MT", "Verdana"

            // Since we're doing testing, include these as well (but not necessary to use results):
            // "Elephant", "Goudy Old Style", "Goudy Stout", "Haettenschweiler", "Harrington", "High Tower Text", "Tahoma"

            int iBitmapWidth = 64;
            int iBitmapHeight = 64;
            // the maximum is 3 * 255 = 765
            // Why 610? Empirical data shows that with this as the check limit, the only exception
            // to the algorithm is Verdana. Have I mentioned I hate calculating Verdana typeface stats?
            // I think Tw Cen MT at 96 DPI has an edge pixel of total RGB of 615 or something, which
            // will throw off the calculations.
            int iColorCheck = 610;

            using (System.Drawing.Bitmap bmGraphics = new System.Drawing.Bitmap(iBitmapWidth, iBitmapHeight))
            {
                // Thanks to John C for submitting this workaround patch!
                // Workaround for Mono: https://bugzilla.xamarin.com/show_bug.cgi?id=33310
                if ((bmGraphics.HorizontalResolution < float.Epsilon) && (bmGraphics.VerticalResolution < float.Epsilon))
                {
                    bmGraphics.SetResolution(96, 96);
                }

                System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmGraphics);

                int i, j;
                bool bFound;
                System.Drawing.Color clr;
                int iPixelStart, iPixelEnd;
                // What's with the double underscore? I learnt it here:
                // http://stackoverflow.com/questions/1833062/how-to-measure-the-pixel-width-of-a-digit-in-a-given-font-size-c/1834064#1834064
                // I call it the Double Underscore Hack, because the original poster didn't name it.
                // It doesn't have to be the underscore character. Just use something distinct enough.
                double fDoubleUnderscoreWidthG = g.MeasureString("__", usablefont).Width;
                double fMaxDigitWidthG = 0;
                double fWidthG = 0;

                // So there are 2 calculation methods used:
                // 1) Directly rendering the digits onto a bitmap and then calculating the width
                //    by determining which pixels are rendered. Since digits (or text) aren't rendered
                //    with a black or white pixel (there's antialiasing), we need a buffer check.
                //    Hence the color check above. I found the sum total of RGB values to be usable.
                // 2) Using the Graphics.MeasureString() function

                // The TextRender.MeasureText() function isn't necessary because the max of the above
                // 2 methods is also greater than or equal to that calculated from MeasureText().
                // Also sidenote: It appears that column widths need an actual rendering
                // (thus the use of the Graphics class, either actual rendering or the use
                // of the MeasureString() function). And that row heights use the
                // TextRender.MeasureText() function. Why is Excel so inconsistent?

                // We have 2 methods because neither of them can definitively determine the maximum
                // digit width for typefaces in 96 and 120 DPI. Even accounting for
                // TextRenderer.MeasureText() (yes, I did empirical experiments with TextRenderer too).
                // It turns out that taking the maximum of the 2 methods work out.
                // Well, except for... wait for it... Verdana.
                // And even then I've only tested them for correctness on a subset of typefaces.
                // Specifically those in the built-in themes. I figure if they're a built-in theme font,
                // people are more likely to use them, so I better make sure they work.

                g.FillRectangle(new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(255, 255, 255)), 0, 0, iBitmapWidth, iBitmapHeight);
                // measure widths of digits 0 to 9
                for (i = 0; i < 10; ++i)
                {
                    // anywhere within the bitmap limits is fine.
                    // (2,2) should be well within a 64 by 64 bitmap
                    g.DrawString(i.ToString(), usablefont, new System.Drawing.SolidBrush(System.Drawing.Color.FromArgb(0, 0, 0)), 2, 2);

                    fWidthG = g.MeasureString(string.Format("_{0}_", i), usablefont).Width - fDoubleUnderscoreWidthG;
                    if (fWidthG > fMaxDigitWidthG)
                    {
                        fMaxDigitWidthG = fWidthG;
                    }
                }

                // For most typefaces, the digit 0 has the largest width. Just for academic interest,
                // Candara has digit 6, Constantia has digit 9, Corbel has digit 6, Impact has digit 6
                // as the digit with the largest width respectively. Yes for both 96 and 120 DPI.
                // Yes, I got them from empirical experimental data.

                iPixelStart = iBitmapWidth;
                iPixelEnd = 0;

                for (j = 0; j < iBitmapWidth; ++j)
                {
                    bFound = false;
                    for (i = 0; i < iBitmapHeight; ++i)
                    {
                        clr = bmGraphics.GetPixel(j, i);
                        if ((clr.R + clr.G + clr.B) < iColorCheck)
                        {
                            bFound = true;
                            break;
                        }
                    }

                    if (bFound)
                    {
                        iPixelStart = j;
                        break;
                    }
                }

                for (j = iBitmapWidth - 1; j >= 0; --j)
                {
                    bFound = false;
                    for (i = 0; i < iBitmapHeight; ++i)
                    {
                        clr = bmGraphics.GetPixel(j, i);
                        if ((clr.R + clr.G + clr.B) < iColorCheck)
                        {
                            bFound = true;
                            break;
                        }
                    }

                    if (bFound)
                    {
                        iPixelEnd = j;
                        break;
                    }
                }

                // +1 because we need to include the start pixel. +1 for an extra pixel buffer.
                // double fMaxDigitWidthR = iPixelEnd - iPixelStart + 1 + 1;
                double fMaxDigitWidthR = iPixelEnd - iPixelStart + 2;

                double fMaxDigitWidthFinal = Math.Max(fMaxDigitWidthR, fMaxDigitWidthG);

                // because Verdana is special. I hate Verdana because all the calculations
                // with the typeface don't make sense.
                if (this.MinorLatinFont.Equals("Verdana", StringComparison.OrdinalIgnoreCase))
                {
                    // this is for 96 DPI
                    if (bmGraphics.HorizontalResolution < 108)
                    {
                        fMaxDigitWidthFinal = 11;
                    }
                    else
                    {
                        fMaxDigitWidthFinal = 12;
                    }
                }

                // Rockwell Condensed and Tw Cen MT Condensed from Office 2013 has exceptions at 96 DPI
                if (this.MinorLatinFont.Equals("Rockwell Condensed", StringComparison.OrdinalIgnoreCase)
                    || this.MinorLatinFont.Equals("Tw Cen MT Condensed", StringComparison.OrdinalIgnoreCase))
                {
                    // this is for 96 DPI
                    if (bmGraphics.HorizontalResolution < 108)
                    {
                        fMaxDigitWidthFinal = 6;
                    }
                }

                fMaxDigitWidthFinal = Math.Ceiling(fMaxDigitWidthFinal);
                int iMaxDigitWidth = Convert.ToInt32(fMaxDigitWidthFinal);

                // basically we're trying to get the closest 1/256 multiple that's less
                // than the pixel interval. Read this article for details and explanations.
                // http://polymathprogrammer.com/2012/11/18/calculate-excel-column-width-pixel-interval/
                listColumnStepSize = new List<double>();
                double fStepInterval = 0;
                double fStepSize = 0.0;
                for (i = 0; i < iMaxDigitWidth; ++i)
                {
                    fStepInterval = (double)i / (double)(iMaxDigitWidth - 1);
                    fStepSize = Math.Truncate(256.0 * fStepInterval) / 256.0;
                    listColumnStepSize.Add(fStepSize);
                }

                // The column width is supposedly in multiples of 8 pixels. Read this for details why:
                // http://support.microsoft.com/kb/214123
                // Besides, Excel supposedly starts with 8 characters as the default.
                // Yes, there are a lot of "supposedly" in the comments. Go ask Microsoft.
                // Hey, multiplying by 8 characters also automatically means the column width
                // is in multiples of 8!
                int iDefaultColumnWidthInPixels = Convert.ToInt32(iMaxDigitWidth * 8);
                int iWholeNumber = iDefaultColumnWidthInPixels / (iMaxDigitWidth - 1);
                int iRemainder = iDefaultColumnWidthInPixels % (iMaxDigitWidth - 1);

                SLDocument.PixelToEMU = Convert.ToInt64((float)SLConstants.InchToEMU / bmGraphics.HorizontalResolution);
                SLDocument.RowHeightMultiple = 72.0 / bmGraphics.VerticalResolution;

                this.iThemeMaxDigitWidth = iMaxDigitWidth;
                this.fThemeColumnWidth = (double)iWholeNumber + listColumnStepSize[iRemainder];
                this.lThemeColumnWidthInEMU = (long)iDefaultColumnWidthInPixels * SLDocument.PixelToEMU;
                this.fThemeRowHeight = SLTool.GetDefaultRowHeight(this.MinorLatinFont);
                this.lThemeRowHeightInEMU = Convert.ToInt64(this.fThemeRowHeight * SLConstants.PointToEMU);
            }
        }

        internal void LoadBuiltinTheme(SLThemeTypeValues themetype)
        {
            switch (themetype)
            {
                case SLThemeTypeValues.Office:
                    this.sThemeName = SLConstants.OfficeThemeName;
                    this.sMajorLatinFont = SLConstants.OfficeThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.OfficeThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.OfficeThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.OfficeThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.OfficeThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.OfficeThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.OfficeThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.OfficeThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.OfficeThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.OfficeThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.OfficeThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.OfficeThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.OfficeThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.OfficeThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Office2013:
                    this.sThemeName = SLConstants.Office2013ThemeName;
                    this.sMajorLatinFont = SLConstants.Office2013ThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.Office2013ThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.Office2013ThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.Office2013ThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.Office2013ThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.Office2013ThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.Office2013ThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.Office2013ThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.Office2013ThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.Office2013ThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.Office2013ThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.Office2013ThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.Office2013ThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.Office2013ThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Adjacency:
                    this.sThemeName = SLConstants.AdjacencyThemeName;
                    this.sMajorLatinFont = SLConstants.AdjacencyThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.AdjacencyThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.AdjacencyThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.AdjacencyThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.AdjacencyThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.AdjacencyThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.AdjacencyThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.AdjacencyThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.AdjacencyThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.AdjacencyThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.AdjacencyThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.AdjacencyThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.AdjacencyThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.AdjacencyThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Angles:
                    this.sThemeName = SLConstants.AnglesThemeName;
                    this.sMajorLatinFont = SLConstants.AnglesThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.AnglesThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.AnglesThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.AnglesThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.AnglesThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.AnglesThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.AnglesThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.AnglesThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.AnglesThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.AnglesThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.AnglesThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.AnglesThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.AnglesThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.AnglesThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Apex:
                    this.sThemeName = SLConstants.ApexThemeName;
                    this.sMajorLatinFont = SLConstants.ApexThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ApexThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ApexThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ApexThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ApexThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ApexThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ApexThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ApexThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ApexThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ApexThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ApexThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ApexThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ApexThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ApexThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Apothecary:
                    this.sThemeName = SLConstants.ApothecaryThemeName;
                    this.sMajorLatinFont = SLConstants.ApothecaryThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ApothecaryThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ApothecaryThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ApothecaryThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ApothecaryThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ApothecaryThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ApothecaryThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ApothecaryThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ApothecaryThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ApothecaryThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ApothecaryThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ApothecaryThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ApothecaryThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ApothecaryThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Aspect:
                    this.sThemeName = SLConstants.AspectThemeName;
                    this.sMajorLatinFont = SLConstants.AspectThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.AspectThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.AspectThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.AspectThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.AspectThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.AspectThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.AspectThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.AspectThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.AspectThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.AspectThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.AspectThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.AspectThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.AspectThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.AspectThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Austin:
                    this.sThemeName = SLConstants.AustinThemeName;
                    this.sMajorLatinFont = SLConstants.AustinThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.AustinThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.AustinThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.AustinThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.AustinThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.AustinThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.AustinThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.AustinThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.AustinThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.AustinThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.AustinThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.AustinThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.AustinThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.AustinThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.BlackTie:
                    this.sThemeName = SLConstants.BlackTieThemeName;
                    this.sMajorLatinFont = SLConstants.BlackTieThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.BlackTieThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.BlackTieThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.BlackTieThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.BlackTieThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.BlackTieThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.BlackTieThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.BlackTieThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.BlackTieThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.BlackTieThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.BlackTieThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.BlackTieThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.BlackTieThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.BlackTieThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Civic:
                    this.sThemeName = SLConstants.CivicThemeName;
                    this.sMajorLatinFont = SLConstants.CivicThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.CivicThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.CivicThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.CivicThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.CivicThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.CivicThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.CivicThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.CivicThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.CivicThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.CivicThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.CivicThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.CivicThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.CivicThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.CivicThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Clarity:
                    this.sThemeName = SLConstants.ClarityThemeName;
                    this.sMajorLatinFont = SLConstants.ClarityThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ClarityThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ClarityThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ClarityThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ClarityThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ClarityThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ClarityThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ClarityThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ClarityThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ClarityThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ClarityThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ClarityThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ClarityThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ClarityThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Composite:
                    this.sThemeName = SLConstants.CompositeThemeName;
                    this.sMajorLatinFont = SLConstants.CompositeThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.CompositeThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.CompositeThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.CompositeThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.CompositeThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.CompositeThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.CompositeThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.CompositeThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.CompositeThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.CompositeThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.CompositeThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.CompositeThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.CompositeThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.CompositeThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Concourse:
                    this.sThemeName = SLConstants.ConcourseThemeName;
                    this.sMajorLatinFont = SLConstants.ConcourseThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ConcourseThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ConcourseThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ConcourseThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ConcourseThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ConcourseThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ConcourseThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ConcourseThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ConcourseThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ConcourseThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ConcourseThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ConcourseThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ConcourseThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ConcourseThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Couture:
                    this.sThemeName = SLConstants.CoutureThemeName;
                    this.sMajorLatinFont = SLConstants.CoutureThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.CoutureThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.CoutureThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.CoutureThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.CoutureThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.CoutureThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.CoutureThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.CoutureThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.CoutureThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.CoutureThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.CoutureThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.CoutureThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.CoutureThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.CoutureThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Elemental:
                    this.sThemeName = SLConstants.ElementalThemeName;
                    this.sMajorLatinFont = SLConstants.ElementalThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ElementalThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ElementalThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ElementalThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ElementalThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ElementalThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ElementalThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ElementalThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ElementalThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ElementalThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ElementalThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ElementalThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ElementalThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ElementalThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Equity:
                    this.sThemeName = SLConstants.EquityThemeName;
                    this.sMajorLatinFont = SLConstants.EquityThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.EquityThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.EquityThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.EquityThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.EquityThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.EquityThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.EquityThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.EquityThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.EquityThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.EquityThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.EquityThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.EquityThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.EquityThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.EquityThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Essential:
                    this.sThemeName = SLConstants.EssentialThemeName;
                    this.sMajorLatinFont = SLConstants.EssentialThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.EssentialThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.EssentialThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.EssentialThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.EssentialThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.EssentialThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.EssentialThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.EssentialThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.EssentialThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.EssentialThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.EssentialThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.EssentialThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.EssentialThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.EssentialThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Executive:
                    this.sThemeName = SLConstants.ExecutiveThemeName;
                    this.sMajorLatinFont = SLConstants.ExecutiveThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ExecutiveThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ExecutiveThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ExecutiveThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ExecutiveThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ExecutiveThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ExecutiveThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ExecutiveThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ExecutiveThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ExecutiveThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ExecutiveThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ExecutiveThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ExecutiveThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ExecutiveThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Facet:
                    this.sThemeName = SLConstants.FacetThemeName;
                    this.sMajorLatinFont = SLConstants.FacetThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.FacetThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.FacetThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.FacetThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.FacetThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.FacetThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.FacetThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.FacetThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.FacetThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.FacetThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.FacetThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.FacetThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.FacetThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.FacetThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Flow:
                    this.sThemeName = SLConstants.FlowThemeName;
                    this.sMajorLatinFont = SLConstants.FlowThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.FlowThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.FlowThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.FlowThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.FlowThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.FlowThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.FlowThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.FlowThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.FlowThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.FlowThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.FlowThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.FlowThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.FlowThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.FlowThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Foundry:
                    this.sThemeName = SLConstants.FoundryThemeName;
                    this.sMajorLatinFont = SLConstants.FoundryThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.FoundryThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.FoundryThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.FoundryThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.FoundryThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.FoundryThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.FoundryThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.FoundryThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.FoundryThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.FoundryThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.FoundryThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.FoundryThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.FoundryThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.FoundryThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Grid:
                    this.sThemeName = SLConstants.GridThemeName;
                    this.sMajorLatinFont = SLConstants.GridThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.GridThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.GridThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.GridThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.GridThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.GridThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.GridThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.GridThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.GridThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.GridThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.GridThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.GridThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.GridThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.GridThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Hardcover:
                    this.sThemeName = SLConstants.HardcoverThemeName;
                    this.sMajorLatinFont = SLConstants.HardcoverThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.HardcoverThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.HardcoverThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.HardcoverThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.HardcoverThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.HardcoverThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.HardcoverThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.HardcoverThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.HardcoverThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.HardcoverThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.HardcoverThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.HardcoverThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.HardcoverThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.HardcoverThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Horizon:
                    this.sThemeName = SLConstants.HorizonThemeName;
                    this.sMajorLatinFont = SLConstants.HorizonThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.HorizonThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.HorizonThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.HorizonThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.HorizonThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.HorizonThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.HorizonThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.HorizonThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.HorizonThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.HorizonThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.HorizonThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.HorizonThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.HorizonThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.HorizonThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Integral:
                    this.sThemeName = SLConstants.IntegralThemeName;
                    this.sMajorLatinFont = SLConstants.IntegralThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.IntegralThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.IntegralThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.IntegralThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.IntegralThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.IntegralThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.IntegralThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.IntegralThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.IntegralThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.IntegralThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.IntegralThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.IntegralThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.IntegralThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.IntegralThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Ion:
                    this.sThemeName = SLConstants.IonThemeName;
                    this.sMajorLatinFont = SLConstants.IonThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.IonThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.IonThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.IonThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.IonThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.IonThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.IonThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.IonThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.IonThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.IonThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.IonThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.IonThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.IonThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.IonThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.IonBoardroom:
                    this.sThemeName = SLConstants.IonBoardroomThemeName;
                    this.sMajorLatinFont = SLConstants.IonBoardroomThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.IonBoardroomThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.IonBoardroomThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.IonBoardroomThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.IonBoardroomThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.IonBoardroomThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.IonBoardroomThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.IonBoardroomThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.IonBoardroomThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.IonBoardroomThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.IonBoardroomThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.IonBoardroomThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.IonBoardroomThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.IonBoardroomThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Median:
                    this.sThemeName = SLConstants.MedianThemeName;
                    this.sMajorLatinFont = SLConstants.MedianThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.MedianThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.MedianThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.MedianThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.MedianThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.MedianThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.MedianThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.MedianThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.MedianThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.MedianThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.MedianThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.MedianThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.MedianThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.MedianThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Metro:
                    this.sThemeName = SLConstants.MetroThemeName;
                    this.sMajorLatinFont = SLConstants.MetroThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.MetroThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.MetroThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.MetroThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.MetroThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.MetroThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.MetroThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.MetroThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.MetroThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.MetroThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.MetroThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.MetroThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.MetroThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.MetroThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Module:
                    this.sThemeName = SLConstants.ModuleThemeName;
                    this.sMajorLatinFont = SLConstants.ModuleThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ModuleThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ModuleThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ModuleThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ModuleThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ModuleThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ModuleThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ModuleThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ModuleThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ModuleThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ModuleThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ModuleThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ModuleThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ModuleThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Newsprint:
                    this.sThemeName = SLConstants.NewsprintThemeName;
                    this.sMajorLatinFont = SLConstants.NewsprintThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.NewsprintThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.NewsprintThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.NewsprintThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.NewsprintThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.NewsprintThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.NewsprintThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.NewsprintThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.NewsprintThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.NewsprintThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.NewsprintThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.NewsprintThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.NewsprintThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.NewsprintThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Opulent:
                    this.sThemeName = SLConstants.OpulentThemeName;
                    this.sMajorLatinFont = SLConstants.OpulentThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.OpulentThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.OpulentThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.OpulentThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.OpulentThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.OpulentThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.OpulentThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.OpulentThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.OpulentThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.OpulentThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.OpulentThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.OpulentThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.OpulentThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.OpulentThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Organic:
                    this.sThemeName = SLConstants.OrganicThemeName;
                    this.sMajorLatinFont = SLConstants.OrganicThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.OrganicThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.OrganicThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.OrganicThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.OrganicThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.OrganicThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.OrganicThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.OrganicThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.OrganicThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.OrganicThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.OrganicThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.OrganicThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.OrganicThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.OrganicThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Oriel:
                    this.sThemeName = SLConstants.OrielThemeName;
                    this.sMajorLatinFont = SLConstants.OrielThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.OrielThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.OrielThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.OrielThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.OrielThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.OrielThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.OrielThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.OrielThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.OrielThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.OrielThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.OrielThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.OrielThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.OrielThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.OrielThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Origin:
                    this.sThemeName = SLConstants.OriginThemeName;
                    this.sMajorLatinFont = SLConstants.OriginThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.OriginThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.OriginThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.OriginThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.OriginThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.OriginThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.OriginThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.OriginThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.OriginThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.OriginThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.OriginThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.OriginThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.OriginThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.OriginThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Paper:
                    this.sThemeName = SLConstants.PaperThemeName;
                    this.sMajorLatinFont = SLConstants.PaperThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.PaperThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.PaperThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.PaperThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.PaperThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.PaperThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.PaperThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.PaperThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.PaperThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.PaperThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.PaperThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.PaperThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.PaperThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.PaperThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Perspective:
                    this.sThemeName = SLConstants.PerspectiveThemeName;
                    this.sMajorLatinFont = SLConstants.PerspectiveThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.PerspectiveThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.PerspectiveThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.PerspectiveThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.PerspectiveThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.PerspectiveThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.PerspectiveThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.PerspectiveThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.PerspectiveThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.PerspectiveThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.PerspectiveThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.PerspectiveThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.PerspectiveThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.PerspectiveThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Pushpin:
                    this.sThemeName = SLConstants.PushpinThemeName;
                    this.sMajorLatinFont = SLConstants.PushpinThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.PushpinThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.PushpinThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.PushpinThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.PushpinThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.PushpinThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.PushpinThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.PushpinThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.PushpinThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.PushpinThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.PushpinThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.PushpinThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.PushpinThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.PushpinThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Retrospect:
                    this.sThemeName = SLConstants.RetrospectThemeName;
                    this.sMajorLatinFont = SLConstants.RetrospectThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.RetrospectThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.RetrospectThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.RetrospectThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.RetrospectThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.RetrospectThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.RetrospectThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.RetrospectThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.RetrospectThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.RetrospectThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.RetrospectThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.RetrospectThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.RetrospectThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.RetrospectThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Slice:
                    this.sThemeName = SLConstants.SliceThemeName;
                    this.sMajorLatinFont = SLConstants.SliceThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.SliceThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.SliceThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.SliceThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.SliceThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.SliceThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.SliceThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.SliceThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.SliceThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.SliceThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.SliceThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.SliceThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.SliceThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.SliceThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Slipstream:
                    this.sThemeName = SLConstants.SlipstreamThemeName;
                    this.sMajorLatinFont = SLConstants.SlipstreamThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.SlipstreamThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.SlipstreamThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.SlipstreamThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.SlipstreamThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.SlipstreamThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.SlipstreamThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.SlipstreamThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.SlipstreamThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.SlipstreamThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.SlipstreamThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.SlipstreamThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.SlipstreamThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.SlipstreamThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Solstice:
                    this.sThemeName = SLConstants.SolsticeThemeName;
                    this.sMajorLatinFont = SLConstants.SolsticeThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.SolsticeThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.SolsticeThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.SolsticeThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.SolsticeThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.SolsticeThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.SolsticeThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.SolsticeThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.SolsticeThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.SolsticeThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.SolsticeThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.SolsticeThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.SolsticeThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.SolsticeThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Technic:
                    this.sThemeName = SLConstants.TechnicThemeName;
                    this.sMajorLatinFont = SLConstants.TechnicThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.TechnicThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.TechnicThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.TechnicThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.TechnicThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.TechnicThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.TechnicThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.TechnicThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.TechnicThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.TechnicThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.TechnicThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.TechnicThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.TechnicThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.TechnicThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Thatch:
                    this.sThemeName = SLConstants.ThatchThemeName;
                    this.sMajorLatinFont = SLConstants.ThatchThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ThatchThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ThatchThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ThatchThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ThatchThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ThatchThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ThatchThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ThatchThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ThatchThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ThatchThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ThatchThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ThatchThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ThatchThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ThatchThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Trek:
                    this.sThemeName = SLConstants.TrekThemeName;
                    this.sMajorLatinFont = SLConstants.TrekThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.TrekThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.TrekThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.TrekThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.TrekThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.TrekThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.TrekThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.TrekThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.TrekThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.TrekThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.TrekThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.TrekThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.TrekThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.TrekThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Urban:
                    this.sThemeName = SLConstants.UrbanThemeName;
                    this.sMajorLatinFont = SLConstants.UrbanThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.UrbanThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.UrbanThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.UrbanThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.UrbanThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.UrbanThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.UrbanThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.UrbanThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.UrbanThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.UrbanThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.UrbanThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.UrbanThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.UrbanThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.UrbanThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Verve:
                    this.sThemeName = SLConstants.VerveThemeName;
                    this.sMajorLatinFont = SLConstants.VerveThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.VerveThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.VerveThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.VerveThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.VerveThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.VerveThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.VerveThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.VerveThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.VerveThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.VerveThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.VerveThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.VerveThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.VerveThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.VerveThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Waveform:
                    this.sThemeName = SLConstants.WaveformThemeName;
                    this.sMajorLatinFont = SLConstants.WaveformThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.WaveformThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.WaveformThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.WaveformThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.WaveformThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.WaveformThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.WaveformThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.WaveformThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.WaveformThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.WaveformThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.WaveformThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.WaveformThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.WaveformThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.WaveformThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Wisp:
                    this.sThemeName = SLConstants.WispThemeName;
                    this.sMajorLatinFont = SLConstants.WispThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.WispThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.WispThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.WispThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.WispThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.WispThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.WispThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.WispThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.WispThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.WispThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.WispThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.WispThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.WispThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.WispThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Autumn:
                    this.sThemeName = SLConstants.AutumnThemeName;
                    this.sMajorLatinFont = SLConstants.AutumnThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.AutumnThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.AutumnThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.AutumnThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.AutumnThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.AutumnThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.AutumnThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.AutumnThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.AutumnThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.AutumnThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.AutumnThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.AutumnThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.AutumnThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.AutumnThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Banded:
                    this.sThemeName = SLConstants.BandedThemeName;
                    this.sMajorLatinFont = SLConstants.BandedThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.BandedThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.BandedThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.BandedThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.BandedThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.BandedThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.BandedThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.BandedThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.BandedThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.BandedThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.BandedThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.BandedThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.BandedThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.BandedThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Basis:
                    this.sThemeName = SLConstants.BasisThemeName;
                    this.sMajorLatinFont = SLConstants.BasisThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.BasisThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.BasisThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.BasisThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.BasisThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.BasisThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.BasisThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.BasisThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.BasisThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.BasisThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.BasisThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.BasisThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.BasisThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.BasisThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Berlin:
                    this.sThemeName = SLConstants.BerlinThemeName;
                    this.sMajorLatinFont = SLConstants.BerlinThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.BerlinThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.BerlinThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.BerlinThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.BerlinThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.BerlinThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.BerlinThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.BerlinThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.BerlinThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.BerlinThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.BerlinThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.BerlinThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.BerlinThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.BerlinThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Celestial:
                    this.sThemeName = SLConstants.CelestialThemeName;
                    this.sMajorLatinFont = SLConstants.CelestialThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.CelestialThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.CelestialThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.CelestialThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.CelestialThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.CelestialThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.CelestialThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.CelestialThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.CelestialThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.CelestialThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.CelestialThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.CelestialThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.CelestialThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.CelestialThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Circuit:
                    this.sThemeName = SLConstants.CircuitThemeName;
                    this.sMajorLatinFont = SLConstants.CircuitThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.CircuitThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.CircuitThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.CircuitThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.CircuitThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.CircuitThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.CircuitThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.CircuitThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.CircuitThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.CircuitThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.CircuitThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.CircuitThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.CircuitThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.CircuitThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Damask:
                    this.sThemeName = SLConstants.DamaskThemeName;
                    this.sMajorLatinFont = SLConstants.DamaskThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.DamaskThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.DamaskThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.DamaskThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.DamaskThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.DamaskThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.DamaskThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.DamaskThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.DamaskThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.DamaskThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.DamaskThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.DamaskThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.DamaskThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.DamaskThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Decatur:
                    this.sThemeName = SLConstants.DecaturThemeName;
                    this.sMajorLatinFont = SLConstants.DecaturThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.DecaturThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.DecaturThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.DecaturThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.DecaturThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.DecaturThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.DecaturThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.DecaturThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.DecaturThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.DecaturThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.DecaturThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.DecaturThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.DecaturThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.DecaturThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Depth:
                    this.sThemeName = SLConstants.DepthThemeName;
                    this.sMajorLatinFont = SLConstants.DepthThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.DepthThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.DepthThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.DepthThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.DepthThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.DepthThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.DepthThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.DepthThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.DepthThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.DepthThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.DepthThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.DepthThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.DepthThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.DepthThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Dividend:
                    this.sThemeName = SLConstants.DividendThemeName;
                    this.sMajorLatinFont = SLConstants.DividendThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.DividendThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.DividendThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.DividendThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.DividendThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.DividendThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.DividendThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.DividendThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.DividendThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.DividendThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.DividendThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.DividendThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.DividendThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.DividendThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Droplet:
                    this.sThemeName = SLConstants.DropletThemeName;
                    this.sMajorLatinFont = SLConstants.DropletThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.DropletThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.DropletThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.DropletThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.DropletThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.DropletThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.DropletThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.DropletThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.DropletThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.DropletThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.DropletThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.DropletThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.DropletThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.DropletThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Frame:
                    this.sThemeName = SLConstants.FrameThemeName;
                    this.sMajorLatinFont = SLConstants.FrameThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.FrameThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.FrameThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.FrameThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.FrameThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.FrameThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.FrameThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.FrameThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.FrameThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.FrameThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.FrameThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.FrameThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.FrameThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.FrameThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Kilter:
                    this.sThemeName = SLConstants.KilterThemeName;
                    this.sMajorLatinFont = SLConstants.KilterThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.KilterThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.KilterThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.KilterThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.KilterThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.KilterThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.KilterThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.KilterThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.KilterThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.KilterThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.KilterThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.KilterThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.KilterThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.KilterThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Macro:
                    this.sThemeName = SLConstants.MacroThemeName;
                    this.sMajorLatinFont = SLConstants.MacroThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.MacroThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.MacroThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.MacroThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.MacroThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.MacroThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.MacroThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.MacroThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.MacroThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.MacroThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.MacroThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.MacroThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.MacroThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.MacroThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.MainEvent:
                    this.sThemeName = SLConstants.MainEventThemeName;
                    this.sMajorLatinFont = SLConstants.MainEventThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.MainEventThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.MainEventThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.MainEventThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.MainEventThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.MainEventThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.MainEventThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.MainEventThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.MainEventThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.MainEventThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.MainEventThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.MainEventThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.MainEventThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.MainEventThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Mesh:
                    this.sThemeName = SLConstants.MeshThemeName;
                    this.sMajorLatinFont = SLConstants.MeshThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.MeshThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.MeshThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.MeshThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.MeshThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.MeshThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.MeshThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.MeshThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.MeshThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.MeshThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.MeshThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.MeshThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.MeshThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.MeshThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Metropolitan:
                    this.sThemeName = SLConstants.MetropolitanThemeName;
                    this.sMajorLatinFont = SLConstants.MetropolitanThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.MetropolitanThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.MetropolitanThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.MetropolitanThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.MetropolitanThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.MetropolitanThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.MetropolitanThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.MetropolitanThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.MetropolitanThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.MetropolitanThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.MetropolitanThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.MetropolitanThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.MetropolitanThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.MetropolitanThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Mylar:
                    this.sThemeName = SLConstants.MylarThemeName;
                    this.sMajorLatinFont = SLConstants.MylarThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.MylarThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.MylarThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.MylarThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.MylarThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.MylarThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.MylarThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.MylarThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.MylarThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.MylarThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.MylarThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.MylarThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.MylarThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.MylarThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Parallax:
                    this.sThemeName = SLConstants.ParallaxThemeName;
                    this.sMajorLatinFont = SLConstants.ParallaxThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ParallaxThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ParallaxThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ParallaxThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ParallaxThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ParallaxThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ParallaxThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ParallaxThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ParallaxThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ParallaxThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ParallaxThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ParallaxThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ParallaxThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ParallaxThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Quotable:
                    this.sThemeName = SLConstants.QuotableThemeName;
                    this.sMajorLatinFont = SLConstants.QuotableThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.QuotableThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.QuotableThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.QuotableThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.QuotableThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.QuotableThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.QuotableThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.QuotableThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.QuotableThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.QuotableThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.QuotableThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.QuotableThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.QuotableThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.QuotableThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Savon:
                    this.sThemeName = SLConstants.SavonThemeName;
                    this.sMajorLatinFont = SLConstants.SavonThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.SavonThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.SavonThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.SavonThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.SavonThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.SavonThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.SavonThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.SavonThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.SavonThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.SavonThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.SavonThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.SavonThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.SavonThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.SavonThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Sketchbook:
                    this.sThemeName = SLConstants.SketchbookThemeName;
                    this.sMajorLatinFont = SLConstants.SketchbookThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.SketchbookThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.SketchbookThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.SketchbookThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.SketchbookThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.SketchbookThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.SketchbookThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.SketchbookThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.SketchbookThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.SketchbookThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.SketchbookThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.SketchbookThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.SketchbookThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.SketchbookThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Slate:
                    this.sThemeName = SLConstants.SlateThemeName;
                    this.sMajorLatinFont = SLConstants.SlateThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.SlateThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.SlateThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.SlateThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.SlateThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.SlateThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.SlateThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.SlateThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.SlateThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.SlateThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.SlateThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.SlateThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.SlateThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.SlateThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Soho:
                    this.sThemeName = SLConstants.SohoThemeName;
                    this.sMajorLatinFont = SLConstants.SohoThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.SohoThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.SohoThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.SohoThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.SohoThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.SohoThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.SohoThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.SohoThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.SohoThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.SohoThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.SohoThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.SohoThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.SohoThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.SohoThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Spring:
                    this.sThemeName = SLConstants.SpringThemeName;
                    this.sMajorLatinFont = SLConstants.SpringThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.SpringThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.SpringThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.SpringThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.SpringThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.SpringThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.SpringThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.SpringThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.SpringThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.SpringThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.SpringThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.SpringThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.SpringThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.SpringThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Summer:
                    this.sThemeName = SLConstants.SummerThemeName;
                    this.sMajorLatinFont = SLConstants.SummerThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.SummerThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.SummerThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.SummerThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.SummerThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.SummerThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.SummerThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.SummerThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.SummerThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.SummerThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.SummerThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.SummerThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.SummerThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.SummerThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Thermal:
                    this.sThemeName = SLConstants.ThermalThemeName;
                    this.sMajorLatinFont = SLConstants.ThermalThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ThermalThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ThermalThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ThermalThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ThermalThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ThermalThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ThermalThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ThermalThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ThermalThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ThermalThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ThermalThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ThermalThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ThermalThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ThermalThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Tradeshow:
                    this.sThemeName = SLConstants.TradeshowThemeName;
                    this.sMajorLatinFont = SLConstants.TradeshowThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.TradeshowThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.TradeshowThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.TradeshowThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.TradeshowThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.TradeshowThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.TradeshowThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.TradeshowThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.TradeshowThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.TradeshowThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.TradeshowThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.TradeshowThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.TradeshowThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.TradeshowThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.UrbanPop:
                    this.sThemeName = SLConstants.UrbanPopThemeName;
                    this.sMajorLatinFont = SLConstants.UrbanPopThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.UrbanPopThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.UrbanPopThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.UrbanPopThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.UrbanPopThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.UrbanPopThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.UrbanPopThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.UrbanPopThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.UrbanPopThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.UrbanPopThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.UrbanPopThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.UrbanPopThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.UrbanPopThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.UrbanPopThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.VaporTrail:
                    this.sThemeName = SLConstants.VaporTrailThemeName;
                    this.sMajorLatinFont = SLConstants.VaporTrailThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.VaporTrailThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.VaporTrailThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.VaporTrailThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.VaporTrailThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.VaporTrailThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.VaporTrailThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.VaporTrailThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.VaporTrailThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.VaporTrailThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.VaporTrailThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.VaporTrailThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.VaporTrailThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.VaporTrailThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.View:
                    this.sThemeName = SLConstants.ViewThemeName;
                    this.sMajorLatinFont = SLConstants.ViewThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.ViewThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.ViewThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.ViewThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.ViewThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.ViewThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.ViewThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.ViewThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.ViewThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.ViewThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.ViewThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.ViewThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.ViewThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.ViewThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.Winter:
                    this.sThemeName = SLConstants.WinterThemeName;
                    this.sMajorLatinFont = SLConstants.WinterThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.WinterThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.WinterThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.WinterThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.WinterThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.WinterThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.WinterThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.WinterThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.WinterThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.WinterThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.WinterThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.WinterThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.WinterThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.WinterThemeFollowedHyperlinkColor;
                    break;
                case SLThemeTypeValues.WoodType:
                    this.sThemeName = SLConstants.WoodTypeThemeName;
                    this.sMajorLatinFont = SLConstants.WoodTypeThemeMajorLatinFont;
                    this.sMinorLatinFont = SLConstants.WoodTypeThemeMinorLatinFont;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark1Color] = SLConstants.WoodTypeThemeDark1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light1Color] = SLConstants.WoodTypeThemeLight1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Dark2Color] = SLConstants.WoodTypeThemeDark2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Light2Color] = SLConstants.WoodTypeThemeLight2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent1Color] = SLConstants.WoodTypeThemeAccent1Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent2Color] = SLConstants.WoodTypeThemeAccent2Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent3Color] = SLConstants.WoodTypeThemeAccent3Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent4Color] = SLConstants.WoodTypeThemeAccent4Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent5Color] = SLConstants.WoodTypeThemeAccent5Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Accent6Color] = SLConstants.WoodTypeThemeAccent6Color;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.Hyperlink] = SLConstants.WoodTypeThemeHyperlink;
                    listThemeColorsHex[(int)SLThemeColorIndexValues.FollowedHyperlinkColor] = SLConstants.WoodTypeThemeFollowedHyperlinkColor;
                    break;
            }

            for (int i = 0; i < listThemeColorsHex.Count; ++i)
            {
                listThemeColors[i] = SLTool.ToColor(listThemeColorsHex[i]);
            }
        }

        internal void LoadTheme(WorkbookPart wbp)
        {
            this.sMajorLatinFont = SLConstants.OfficeThemeMajorLatinFont;
            this.sMinorLatinFont = SLConstants.OfficeThemeMinorLatinFont;

            System.Drawing.Color clr = new System.Drawing.Color();
            int index = 0;
            if (wbp.ThemePart != null)
            {
                OpenXmlReader oxr = OpenXmlReader.Create(wbp.ThemePart);
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(A.Dark1Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Dark1Color;
                        A.Dark1Color dk1 = (A.Dark1Color)oxr.LoadCurrentElement();
                        if (dk1.RgbColorModelHex != null && dk1.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(dk1.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (dk1.SystemColor != null && dk1.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(dk1.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Light1Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Light1Color;
                        A.Light1Color lt1 = (A.Light1Color)oxr.LoadCurrentElement();
                        if (lt1.RgbColorModelHex != null && lt1.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(lt1.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (lt1.SystemColor != null && lt1.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(lt1.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Dark2Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Dark2Color;
                        A.Dark2Color dk2 = (A.Dark2Color)oxr.LoadCurrentElement();
                        if (dk2.RgbColorModelHex != null && dk2.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(dk2.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (dk2.SystemColor != null && dk2.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(dk2.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Light2Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Light2Color;
                        A.Light2Color lt2 = (A.Light2Color)oxr.LoadCurrentElement();
                        if (lt2.RgbColorModelHex != null && lt2.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(lt2.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (lt2.SystemColor != null && lt2.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(lt2.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Accent1Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Accent1Color;
                        A.Accent1Color accent1 = (A.Accent1Color)oxr.LoadCurrentElement();
                        if (accent1.RgbColorModelHex != null && accent1.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(accent1.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (accent1.SystemColor != null && accent1.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(accent1.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Accent2Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Accent2Color;
                        A.Accent2Color accent2 = (A.Accent2Color)oxr.LoadCurrentElement();
                        if (accent2.RgbColorModelHex != null && accent2.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(accent2.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (accent2.SystemColor != null && accent2.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(accent2.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Accent3Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Accent3Color;
                        A.Accent3Color accent3 = (A.Accent3Color)oxr.LoadCurrentElement();
                        if (accent3.RgbColorModelHex != null && accent3.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(accent3.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (accent3.SystemColor != null && accent3.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(accent3.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Accent4Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Accent4Color;
                        A.Accent4Color accent4 = (A.Accent4Color)oxr.LoadCurrentElement();
                        if (accent4.RgbColorModelHex != null && accent4.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(accent4.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (accent4.SystemColor != null && accent4.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(accent4.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Accent5Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Accent5Color;
                        A.Accent5Color accent5 = (A.Accent5Color)oxr.LoadCurrentElement();
                        if (accent5.RgbColorModelHex != null && accent5.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(accent5.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (accent5.SystemColor != null && accent5.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(accent5.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Accent6Color))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Accent6Color;
                        A.Accent6Color accent6 = (A.Accent6Color)oxr.LoadCurrentElement();
                        if (accent6.RgbColorModelHex != null && accent6.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(accent6.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (accent6.SystemColor != null && accent6.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(accent6.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.Hyperlink))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.Hyperlink;
                        A.Hyperlink hlink = (A.Hyperlink)oxr.LoadCurrentElement();
                        if (hlink.RgbColorModelHex != null && hlink.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(hlink.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (hlink.SystemColor != null && hlink.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(hlink.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.FollowedHyperlinkColor))
                    {
                        clr = new System.Drawing.Color();
                        index = (int)SLThemeColorIndexValues.FollowedHyperlinkColor;
                        A.FollowedHyperlinkColor fhlink = (A.FollowedHyperlinkColor)oxr.LoadCurrentElement();
                        if (fhlink.RgbColorModelHex != null && fhlink.RgbColorModelHex.Val != null)
                        {
                            clr = SLTool.ToColor(fhlink.RgbColorModelHex.Val);
                            listThemeColors[index] = clr;
                        }
                        else if (fhlink.SystemColor != null && fhlink.SystemColor.LastColor != null)
                        {
                            clr = SLTool.ToColor(fhlink.SystemColor.LastColor.Value);
                            listThemeColors[index] = clr;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.MajorFont))
                    {
                        A.MajorFont major = (A.MajorFont)oxr.LoadCurrentElement();
                        if (major.LatinFont != null && major.LatinFont.Typeface != null)
                        {
                            this.sMajorLatinFont = major.LatinFont.Typeface.Value;
                        }
                    }
                    else if (oxr.ElementType == typeof(A.MinorFont))
                    {
                        A.MinorFont minor = (A.MinorFont)oxr.LoadCurrentElement();
                        if (minor.LatinFont != null && minor.LatinFont.Typeface != null)
                        {
                            this.sMinorLatinFont = minor.LatinFont.Typeface.Value;
                        }
                    }
                }
                oxr.Dispose();
            }
        }

        // NOTE: indexed colours are for supporting legacy spreadsheets.
        internal void LoadIndexedColors(WorkbookPart wbp)
        {
            listIndexedColors = new List<System.Drawing.Color>();

            IndexedColors ic;
            RgbColor rgbclr;
            bool bHasIndexedColors = false;

            if (wbp.WorkbookStylesPart != null)
            {
                OpenXmlReader oxr = OpenXmlReader.Create(wbp.WorkbookStylesPart);
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(IndexedColors))
                    {
                        bHasIndexedColors = true;
                        ic = (IndexedColors)oxr.LoadCurrentElement();
                        OpenXmlReader oxrIndexed = OpenXmlReader.Create(ic);
                        while (oxrIndexed.Read())
                        {
                            if (oxrIndexed.ElementType == typeof(RgbColor))
                            {
                                rgbclr = (RgbColor)oxrIndexed.LoadCurrentElement();
                                if (rgbclr.Rgb != null)
                                {
                                    listIndexedColors.Add(SLTool.ToColor(rgbclr.Rgb.Value));
                                }
                                else
                                {
                                    listIndexedColors.Add(new System.Drawing.Color());
                                }
                            }
                        }
                        oxrIndexed.Dispose();

                        break;
                    }
                }
                oxr.Dispose();
            }

            // if there are no indexed colours, load in a default palette
            if (!bHasIndexedColors)
            {
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0xFF, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0xFF, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0, 0));

                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0xFF, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0xFF, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x80, 0, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0x80, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0, 0x80));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x80, 0x80, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x80, 0, 0x80));

                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0x80, 0x80));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xC0, 0xC0, 0xC0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x80, 0x80, 0x80));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x99, 0x99, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x99, 0x33, 0x66));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xCC));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xCC, 0xFF, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x66, 0, 0x66));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0x80, 0x80));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0x66, 0xCC));

                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xCC, 0xCC, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0, 0x80));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0xFF, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x80, 0, 0x80));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x80, 0, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0x80, 0x80));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0xCC, 0xFF));

                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xCC, 0xFF, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xCC, 0xFF, 0xCC));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x99));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x99, 0xCC, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0x99, 0xCC));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xCC, 0x99, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xCC, 0x99));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x33, 0x66, 0xFF));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x33, 0xCC, 0xCC));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x99, 0xCC, 0));

                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xCC, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0x99, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0x66, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x66, 0x66, 0x99));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x96, 0x96, 0x96));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0x33, 0x66));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x33, 0x99, 0x66));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0, 0x33, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x33, 0x33, 0));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x99, 0x33, 0));

                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x99, 0x33, 0x66));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x33, 0x33, 0x99));
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x33, 0x33, 0x33));

                // index 64. Don't know the system foreground color, so just use black
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0x00, 0x00, 0x00));
                // index 65. Don't know the system background color, so just use white
                listIndexedColors.Add(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xFF));
            }
        }
    }
}
