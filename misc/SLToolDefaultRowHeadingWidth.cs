using System;

namespace SpreadsheetLight
{
    internal partial class SLTool
    {
        /// <summary>
        /// This returns the width in points
        /// </summary>
        /// <param name="FontName"></param>
        /// <param name="ThrowExceptionsIfAny"></param>
        /// <returns></returns>
        internal static double GetDefaultRowHeadingWidth(string FontName, bool ThrowExceptionsIfAny)
        {
            // Yes, this is a long list of hard-coded values. Blame Excel, because I don't
            // know how to calculate that top-left corner box that's to the left of column
            // headings and above the row headings, and depending on which DPI your screen is in.

            // I'm not doing the full list of available fonts on my computer because although
            // I'm dedicated to making sure the values are correct, I'm not insane. At least,
            // not yet. Maybe some drearily boring raining afternoon where watching paint dry
            // is not satisfactorily filling up my time, I'll do the full list.
            // I'll just handle the theme fonts, on the basis that most people will use the theme
            // fonts more often.
            // There's a default case statement, but don't bet the farm on the algorithm...

            // How are the case statements generated? I took the list of theme fonts and ran with it.
            // Generated 2 sets of spreadsheets, one for 96 DPI and one for 120 DPI. I'm too lazy
            // to do for 144 DPI as well.
            // I have 2 worksheets in each spreadsheet. Then I split one worksheet at E2 (doesn't
            // matter where, but just need to take note). Then I went to the other worksheet and
            // set the headings checkbox to false. Then I split this worksheet at the same position
            // (E2 in my case). Then I save the spreadsheet.

            // Then I wrote another small program to read in all the spreadsheets and you'll get
            // something like these for the 2 worksheets:

            //  <x:sheetViews>
            //    <x:sheetView workbookViewId="0">
            //      <x:pane xSplit="4248" ySplit="576" topLeftCell="E2" activePane="bottomRight" />
            //      <x:selection pane="topRight" activeCell="E1" sqref="E1" />
            //      <x:selection pane="bottomLeft" activeCell="A2" sqref="A2" />
            //      <x:selection pane="bottomRight" activeCell="E2" sqref="E2" />
            //    </x:sheetView>
            //  </x:sheetViews>

            //  <x:sheetViews>
            //    <x:sheetView showRowColHeaders="0" tabSelected="1" workbookViewId="0">
            //      <x:pane xSplit="3840" ySplit="288" topLeftCell="E2" activePane="bottomRight" />
            //      <x:selection pane="topRight" activeCell="E1" sqref="E1" />
            //      <x:selection pane="bottomLeft" activeCell="A2" sqref="A2" />
            //      <x:selection pane="bottomRight" activeCell="E2" sqref="E2" />
            //    </x:sheetView>
            //  </x:sheetViews>

            // The width of that pesky box is the difference between the xSplit values.
            // In this case, it's 4248 - 3840 = 408. This value is in twentieths of a point,
            // so the final value is 408 / 20 = 20.4, and this is for Calibri 120 DPI.

            string fontcheck = string.Empty;
            double fDefaultRowHeadingWidth = 20.4;
            float fResolution = 96;
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(32, 32);
            fResolution = bm.VerticalResolution;
            bm.Dispose();

            //96,120
            if (fResolution < 108)
            {
                fDefaultRowHeadingWidth = 19.5;
                fontcheck = string.Format("{0} [96]", FontName);
            }
            else
            {
                fDefaultRowHeadingWidth = 20.4;
                fontcheck = string.Format("{0} [120]", FontName);
            }

            switch (fontcheck)
            {
                case "Arial [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Arial [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Arial Black [96]":
                    fDefaultRowHeadingWidth = 27.75;
                    break;
                case "Arial Black [120]":
                    fDefaultRowHeadingWidth = 25.8;
                    break;
                case "Arial Narrow [96]":
                    fDefaultRowHeadingWidth = 19.5;
                    break;
                case "Arial Narrow [120]":
                    fDefaultRowHeadingWidth = 17.4;
                    break;
                case "Bodoni MT Condensed [96]":
                    fDefaultRowHeadingWidth = 17.25;
                    break;
                case "Bodoni MT Condensed [120]":
                    fDefaultRowHeadingWidth = 15.6;
                    break;
                case "Book Antiqua [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Book Antiqua [120]":
                    fDefaultRowHeadingWidth = 20.4;
                    break;
                case "Bookman Old Style [96]":
                    fDefaultRowHeadingWidth = 27.75;
                    break;
                case "Bookman Old Style [120]":
                    fDefaultRowHeadingWidth = 25.8;
                    break;
                case "Calibri Light [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Calibri Light [120]":
                    fDefaultRowHeadingWidth = 20.4;
                    break;
                case "Calibri [96]":
                    fDefaultRowHeadingWidth = 19.5;
                    break;
                case "Calibri [120]":
                    fDefaultRowHeadingWidth = 20.4;
                    break;
                case "Calisto MT [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Calisto MT [120]":
                    fDefaultRowHeadingWidth = 20.4;
                    break;
                case "Cambria [96]":
                    fDefaultRowHeadingWidth = 25.5;
                    break;
                case "Cambria [120]":
                    fDefaultRowHeadingWidth = 24;
                    break;
                case "Candara [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Candara [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Century Gothic [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Century Gothic [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Century Schoolbook [96]":
                    fDefaultRowHeadingWidth = 25.5;
                    break;
                case "Century Schoolbook [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Consolas [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Consolas [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Constantia [96]":
                    fDefaultRowHeadingWidth = 25.5;
                    break;
                case "Constantia [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Corbel [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Corbel [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Franklin Gothic Book [96]":
                    fDefaultRowHeadingWidth = 27.75;
                    break;
                case "Franklin Gothic Book [120]":
                    fDefaultRowHeadingWidth = 25.8;
                    break;
                case "Franklin Gothic Medium [96]":
                    fDefaultRowHeadingWidth = 27.75;
                    break;
                case "Franklin Gothic Medium [120]":
                    fDefaultRowHeadingWidth = 25.8;
                    break;
                case "Garamond [96]":
                    fDefaultRowHeadingWidth = 19.5;
                    break;
                case "Garamond [120]":
                    fDefaultRowHeadingWidth = 17.4;
                    break;
                case "Georgia [96]":
                    fDefaultRowHeadingWidth = 27.75;
                    break;
                case "Georgia [120]":
                    fDefaultRowHeadingWidth = 28.8;
                    break;
                case "Gill Sans MT [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Gill Sans MT [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Impact [96]":
                    fDefaultRowHeadingWidth = 25.5;
                    break;
                case "Impact [120]":
                    fDefaultRowHeadingWidth = 24;
                    break;
                case "Lucida Sans [96]":
                    fDefaultRowHeadingWidth = 27.75;
                    break;
                case "Lucida Sans [120]":
                    fDefaultRowHeadingWidth = 25.8;
                    break;
                case "Lucida Sans Unicode [96]":
                    fDefaultRowHeadingWidth = 27.75;
                    break;
                case "Lucida Sans Unicode [120]":
                    fDefaultRowHeadingWidth = 25.8;
                    break;
                case "Palatino Linotype [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Palatino Linotype [120]":
                    fDefaultRowHeadingWidth = 20.4;
                    break;
                case "Perpetua [96]":
                    fDefaultRowHeadingWidth = 19.5;
                    break;
                case "Perpetua [120]":
                    fDefaultRowHeadingWidth = 17.4;
                    break;
                case "Rockwell Condensed [96]":
                    fDefaultRowHeadingWidth = 15;
                    break;
                case "Rockwell Condensed [120]":
                    fDefaultRowHeadingWidth = 15.6;
                    break;
                case "Rockwell [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Rockwell [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Times New Roman [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Times New Roman [120]":
                    fDefaultRowHeadingWidth = 20.4;
                    break;
                case "Trebuchet MS [96]":
                    fDefaultRowHeadingWidth = 25.5;
                    break;
                case "Trebuchet MS [120]":
                    fDefaultRowHeadingWidth = 24;
                    break;
                case "Tw Cen MT Condensed [96]":
                    fDefaultRowHeadingWidth = 15;
                    break;
                case "Tw Cen MT Condensed [120]":
                    fDefaultRowHeadingWidth = 15.6;
                    break;
                case "Tw Cen MT [96]":
                    fDefaultRowHeadingWidth = 21.75;
                    break;
                case "Tw Cen MT [120]":
                    fDefaultRowHeadingWidth = 22.2;
                    break;
                case "Verdana [96]":
                    fDefaultRowHeadingWidth = 30;
                    break;
                case "Verdana [120]":
                    fDefaultRowHeadingWidth = 28.8;
                    break;
                default:
                    // I'm going to try to figure out something out for this... no guarantees.
                    double fRatio = (double)fResolution / 96.0;
                    // Apparently, the row headings are in font size 11 when in 96 DPI,
                    // and in font size 14 when in 120 DPI. (go do a screenshot of Excel and see
                    // for yourself).
                    // 11 pt * 120 DPI / 96 DPI = 13.75 pt, but we deal with at least
                    // half point font sizes, so we round to 14 pt. I think. Who knows
                    // what Excel is really doing... I'm not going to do "round to nearest half point",
                    // because there's really no point... I don't even know if this makes
                    // any difference at all...
                    float fFontSize = (float)Math.Round(SLConstants.DefaultFontSize * fRatio);
                    double fPadding = Math.Ceiling(5 * fRatio);
                    // Because the row heading is about 3 digits wide. Why 300? At full screen, the
                    // bottom row is around 30, and we need 3 digits, so I added a zero at the end.
                    // Of course, having a movie title is cool too... "This is Sparta!" *kicks Excel*
                    string sText = "300";
                    using (System.Drawing.Bitmap bmExtra = new System.Drawing.Bitmap(128, 128))
                    {
                        using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bmExtra))
                        {
                            fDefaultRowHeadingWidth = (double)SLTool.MeasureText(bmExtra, g, sText, SLTool.GetUsableNormalFont(FontName, fFontSize, System.Drawing.FontStyle.Regular, ThrowExceptionsIfAny)).Width;
                            fDefaultRowHeadingWidth += fPadding;
                            fDefaultRowHeadingWidth = fDefaultRowHeadingWidth * (double)SLDocument.PixelToEMU / (double)SLConstants.PointToEMU;
                        }
                    }
                    break;
            }

            return fDefaultRowHeadingWidth;
        }
    }
}
