using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// Indicates if the row has an existing style.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>True if the row has an existing style. False otherwise.</returns>
        public bool HasRowStyle(int RowIndex)
        {
            bool result = false;
            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                SLRowProperties rp = slws.RowProperties[RowIndex];
                if (rp.StyleIndex > 0)
                {
                    result = true;
                }
            }

            return result;
        }

        /// <summary>
        /// Get the row height. If the row doesn't have a height explicitly set, the default row height for the current worksheet is returned.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>The row height in points.</returns>
        public double GetRowHeight(int RowIndex)
        {
            double fHeight = slws.SheetFormatProperties.DefaultRowHeight;
            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                SLRowProperties rp = slws.RowProperties[RowIndex];
                if (rp.HasHeight)
                {
                    fHeight = rp.Height;
                }
            }

            return fHeight;
        }

        /// <summary>
        /// Set the row height.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="RowHeight">The row height in points.</param>
        /// <returns>True if the row index is valid. False otherwise.</returns>
        public bool SetRowHeight(int RowIndex, double RowHeight)
        {
            return SetRowHeight(RowIndex, RowIndex, RowHeight);
        }

        /// <summary>
        /// Set the row height for a range of rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <param name="RowHeight">The row height in points.</param>
        /// <returns>True if the row indices are valid. False otherwise.</returns>
        public bool SetRowHeight(int StartRowIndex, int EndRowIndex, double RowHeight)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            bool result = false;
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                result = true;
                int i = 0;
                SLRowProperties rp;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        rp.Height = RowHeight;
                        slws.RowProperties[i] = rp;
                    }
                    else
                    {
                        rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                        rp.Height = RowHeight;
                        slws.RowProperties.Add(i, rp);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Automatically fit row height according to cell contents.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        public void AutoFitRow(int RowIndex)
        {
            this.AutoFitRow(RowIndex, RowIndex, SLConstants.MaximumRowHeight);
        }

        /// <summary>
        /// Automatically fit row height according to cell contents.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="MaximumRowHeight">The maximum row height in points.</param>
        public void AutoFitRow(int RowIndex, double MaximumRowHeight)
        {
            this.AutoFitRow(RowIndex, RowIndex, MaximumRowHeight);
        }

        /// <summary>
        /// Automatically fit row height according to cell contents.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        public void AutoFitRow(int StartRowIndex, int EndRowIndex)
        {
            this.AutoFitRow(StartRowIndex, EndRowIndex, SLConstants.MaximumRowHeight);
        }

        /// <summary>
        /// Automatically fit row height according to cell contents.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <param name="MaximumRowHeight">The maximum row height in points.</param>
        public void AutoFitRow(int StartRowIndex, int EndRowIndex, double MaximumRowHeight)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iStartRowIndex > SLConstants.RowLimit) iStartRowIndex = SLConstants.RowLimit;
            if (iEndRowIndex < 1) iEndRowIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;

            if (MaximumRowHeight > SLConstants.MaximumRowHeight) MaximumRowHeight = SLConstants.MaximumRowHeight;
            int iMaximumPixelLength = Convert.ToInt32(Math.Floor(MaximumRowHeight / SLDocument.RowHeightMultiple));

            Dictionary<int, int> pixellength = this.AutoFitRowColumn(true, iStartRowIndex, iEndRowIndex, iMaximumPixelLength);

            double fDefaultRowHeight = slws.SheetFormatProperties.DefaultRowHeight;
            double fMinimumHeight = 0;

            SLStyle style;
            int iStyleIndex;
            string sFontName;
            double fFontSize;
            bool bBold;
            bool bItalic;
            bool bStrike;
            bool bUnderline;
            System.Drawing.FontStyle drawstyle = System.Drawing.FontStyle.Regular;
            System.Drawing.Font ftUsableFont;
            System.Drawing.SizeF szf;

            using (System.Drawing.Bitmap bm = new System.Drawing.Bitmap(4096, 2048))
            {
                double fResolution = 96.0;
                fResolution = (double)bm.VerticalResolution;

                using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bm))
                {
                    SLRowProperties rp;
                    double fRowHeight;
                    int iPixelLength;
                    foreach (int pixlenpt in pixellength.Keys)
                    {
                        iPixelLength = pixellength[pixlenpt];
                        if (iPixelLength > 0)
                        {
                            // height in points = number of pixels * 72 (points per inch) / resolution (DPI)
                            fRowHeight = (double)iPixelLength * 72.0 / fResolution;
                            if (slws.RowProperties.ContainsKey(pixlenpt))
                            {
                                rp = slws.RowProperties[pixlenpt];

                                iStyleIndex = (int)rp.StyleIndex;
                                if (dictAutoFitFontCache.ContainsKey(iStyleIndex))
                                {
                                    ftUsableFont = dictAutoFitFontCache[iStyleIndex];
                                }
                                else
                                {
                                    style = new SLStyle();
                                    style.FromHash(listStyle[iStyleIndex]);

                                    sFontName = SimpleTheme.MinorLatinFont;
                                    fFontSize = SLConstants.DefaultFontSize;
                                    bBold = false;
                                    bItalic = false;
                                    bStrike = false;
                                    bUnderline = false;
                                    drawstyle = System.Drawing.FontStyle.Regular;
                                    if (style.fontReal.HasFontScheme)
                                    {
                                        if (style.fontReal.FontScheme == FontSchemeValues.Major) sFontName = SimpleTheme.MajorLatinFont;
                                        else if (style.fontReal.FontScheme == FontSchemeValues.Minor) sFontName = SimpleTheme.MinorLatinFont;
                                        else if (style.fontReal.FontName != null && style.fontReal.FontName.Length > 0) sFontName = style.fontReal.FontName;
                                    }
                                    else
                                    {
                                        if (style.fontReal.FontName != null && style.fontReal.FontName.Length > 0) sFontName = style.fontReal.FontName;
                                    }

                                    if (style.fontReal.FontSize != null) fFontSize = style.fontReal.FontSize.Value;
                                    if (style.fontReal.Bold != null && style.fontReal.Bold.Value) bBold = true;
                                    if (style.fontReal.Italic != null && style.fontReal.Italic.Value) bItalic = true;
                                    if (style.fontReal.Strike != null && style.fontReal.Strike.Value) bStrike = true;
                                    if (style.fontReal.HasUnderline) bUnderline = true;

                                    if (bBold) drawstyle |= System.Drawing.FontStyle.Bold;
                                    if (bItalic) drawstyle |= System.Drawing.FontStyle.Italic;
                                    if (bStrike) drawstyle |= System.Drawing.FontStyle.Strikeout;
                                    if (bUnderline) drawstyle |= System.Drawing.FontStyle.Underline;

                                    ftUsableFont = SLTool.GetUsableNormalFont(sFontName, fFontSize, drawstyle, gbThrowExceptionsIfAny);

                                    dictAutoFitFontCache[iStyleIndex] = (System.Drawing.Font)ftUsableFont.Clone();
                                }

                                // if the row has a style, we have to check if the resulting row height
                                // based on the typeface, boldness, italicise-ness and whatnot will change.
                                // Basically we're calculating a "default" row height based on the typeface
                                // set on the entire row.
                                // any text will do. Apparently the height is the same regardless.
                                szf = SLTool.MeasureText(bm, g, "0123456789", ftUsableFont);
                                fMinimumHeight = Math.Min(szf.Height, fDefaultRowHeight);

                                if (fRowHeight < fMinimumHeight)
                                {
                                    rp.Height = fMinimumHeight;
                                    rp.CustomHeight = false;
                                    slws.RowProperties[pixlenpt] = rp.Clone();
                                }
                            }
                            else
                            {
                                rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                                rp.Height = fRowHeight;
                                rp.CustomHeight = false;
                                slws.RowProperties[pixlenpt] = rp.Clone();
                            }
                        }
                        else
                        {
                            // else we set autoheight. Meaning we set the default height for any
                            // existing rows.

                            if (slws.RowProperties.ContainsKey(pixlenpt))
                            {
                                rp = slws.RowProperties[pixlenpt];
                                rp.Height = SimpleTheme.ThemeRowHeight;
                                rp.CustomHeight = false;
                                slws.RowProperties[pixlenpt] = rp.Clone();
                            }
                        }
                    }

                    // end of Graphics
                }
            }
        }

        /// <summary>
        /// Indicates if the row is hidden.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>True if the row is hidden. False otherwise.</returns>
        public bool IsRowHidden(int RowIndex)
        {
            bool result = false;
            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                SLRowProperties rp = slws.RowProperties[RowIndex];
                result = rp.Hidden;
            }

            return result;
        }

        private bool ToggleRowHidden(int StartRowIndex, int EndRowIndex, bool Hidden)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            bool result = false;
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                result = true;
                int i = 0;
                SLRowProperties rp;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        rp.Hidden = Hidden;
                        slws.RowProperties[i] = rp;
                    }
                    else
                    {
                        rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                        rp.Hidden = Hidden;
                        slws.RowProperties.Add(i, rp);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Hide the row.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>True if the row index is valid. False otherwise.</returns>
        public bool HideRow(int RowIndex)
        {
            return ToggleRowHidden(RowIndex, RowIndex, true);
        }

        /// <summary>
        /// Hide a range of rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <returns>True if the row indices are valid. False otherwise.</returns>
        public bool HideRow(int StartRowIndex, int EndRowIndex)
        {
            return ToggleRowHidden(StartRowIndex, EndRowIndex, true);
        }

        /// <summary>
        /// Unhide the row.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>True if the row index is valid. False otherwise.</returns>
        public bool UnhideRow(int RowIndex)
        {
            return ToggleRowHidden(RowIndex, RowIndex, false);
        }

        /// <summary>
        /// Unhide a range of rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <returns>True if the row indices are valid. False otherwise.</returns>
        public bool UnhideRow(int StartRowIndex, int EndRowIndex)
        {
            return ToggleRowHidden(StartRowIndex, EndRowIndex, false);
        }

        /// <summary>
        /// Get the group level of the row.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>The group level (between 0 and 7, both inclusive). The default is 0.</returns>
        public int GetRowGroupLevel(int RowIndex)
        {
            int iGroupLevel = 0;
            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                iGroupLevel = slws.RowProperties[RowIndex].OutlineLevel;
            }

            return iGroupLevel;
        }

        /// <summary>
        /// Group rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row.</param>
        /// <param name="EndRowIndex">The row index of the end row.</param>
        public void GroupRows(int StartRowIndex, int EndRowIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            // I haven't personally checked this, but there's a collapsing -/+ box on the row
            // just below the grouped rows. This implies the very last row that can be
            // grouped is the (row limit - 1)th row, because (row limit)th row will have that
            // collapsing box.
            if (iEndRowIndex >= SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit - 1;
            // there's nothing to group...
            if (iStartRowIndex > iEndRowIndex) return;

            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                int i = 0;
                SLRowProperties rp;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        // Excel supports only levels 0 to 7
                        if (rp.OutlineLevel < 7) ++rp.OutlineLevel;
                        slws.RowProperties[i] = rp.Clone();
                    }
                    else
                    {
                        rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                        rp.OutlineLevel = 1;
                        slws.RowProperties[i] = rp.Clone();
                    }
                }
            }
        }

        /// <summary>
        /// Ungroup rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row.</param>
        /// <param name="EndRowIndex">The row index of the end row.</param>
        public void UngroupRows(int StartRowIndex, int EndRowIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            // the following algorithm is not guaranteed to work in all cases.
            // The data is sort of loosely linked together with no guarantee that they
            // all make sense together. If you use Excel, then the internal data is sort of
            // guaranteed to make sense together, but anyone can make an Open XML spreadsheet
            // with possibly invalid-looking data. Maybe Excel will accept it, maybe not.

            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                SLRowProperties rp;
                byte byCurrentOutlineLevel;
                int i;

                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        if (rp.OutlineLevel > 0) --rp.OutlineLevel;
                        slws.RowProperties[i] = rp.Clone();

                        // if after ungrouping, the outline level is the same as the next
                        // one and the next one is collapsed, then we probably reached the
                        // end of the group and we uncollapse the thing. It's not so much
                        // an uncollapse but an indication to tell the application
                        // (read: Excel) not to choke on missing groups with a collapse command.
                        byCurrentOutlineLevel = rp.OutlineLevel;
                        if (slws.RowProperties.ContainsKey(i + 1))
                        {
                            rp = slws.RowProperties[i + 1];
                            if (rp.OutlineLevel == byCurrentOutlineLevel && rp.Collapsed)
                            {
                                rp.Collapsed = false;
                                slws.RowProperties[i + 1] = rp.Clone();
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Collapse a group of rows.
        /// </summary>
        /// <param name="RowIndex">The row index of the row just after the group of rows you want to collapse. For example, this will be row 5 if rows 2 to 4 are grouped.</param>
        public void CollapseRows(int RowIndex)
        {
            if (RowIndex < 1 || RowIndex > SLConstants.RowLimit) return;

            // the following algorithm is not guaranteed to work in all cases.
            // The data is sort of loosely linked together with no guarantee that they
            // all make sense together. If you use Excel, then the internal data is sort of
            // guaranteed to make sense together, but anyone can make an Open XML spreadsheet
            // with possibly invalid-looking data. Maybe Excel will accept it, maybe not.

            SLRowProperties rp;
            byte byCurrentOutlineLevel = 0;
            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                rp = slws.RowProperties[RowIndex];
                byCurrentOutlineLevel = rp.OutlineLevel;
            }

            bool bFound = false;
            int i;

            for (i = RowIndex - 1; i >= 1; --i)
            {
                if (slws.RowProperties.ContainsKey(i))
                {
                    rp = slws.RowProperties[i];
                    if (rp.OutlineLevel > byCurrentOutlineLevel)
                    {
                        bFound = true;
                        rp.Hidden = true;
                        slws.RowProperties[i] = rp.Clone();
                    }
                    else break;
                }
                else break;
            }

            if (bFound)
            {
                if (slws.RowProperties.ContainsKey(RowIndex))
                {
                    rp = slws.RowProperties[RowIndex];
                    rp.Collapsed = true;
                    slws.RowProperties[RowIndex] = rp.Clone();
                }
                else
                {
                    rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                    rp.Collapsed = true;
                    slws.RowProperties[RowIndex] = rp.Clone();
                }
            }
        }

        /// <summary>
        /// Expand a group of rows.
        /// </summary>
        /// <param name="RowIndex">The row index of the row just after the group of rows you want to expand. For example, this will be row 5 if rows 2 to 4 are grouped.</param>
        public void ExpandRows(int RowIndex)
        {
            if (RowIndex < 1 || RowIndex > SLConstants.RowLimit) return;

            // the following algorithm is not guaranteed to work in all cases.
            // The data is sort of loosely linked together with no guarantee that they
            // all make sense together. If you use Excel, then the internal data is sort of
            // guaranteed to make sense together, but anyone can make an Open XML spreadsheet
            // with possibly invalid-looking data. Maybe Excel will accept it, maybe not.

            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                SLRowProperties rp = slws.RowProperties[RowIndex];
                // no point if it's not the collapsing -/+ box
                if (rp.Collapsed)
                {
                    if (rp.Hidden)
                    {
                        // if it's hidden, it's probably because it and it's associated
                        // group is hidden behind another group. So we don't show the rest
                        // of the group.
                        rp.Collapsed = false;
                        slws.RowProperties[RowIndex] = rp.Clone();
                        // Of course I don't really know that for sure. Hence the "probably".
                    }
                    else
                    {
                        rp.Collapsed = false;
                        slws.RowProperties[RowIndex] = rp.Clone();

                        byte byCurrentOutlineLevel = rp.OutlineLevel;
                        int i;
                        for (i = RowIndex - 1; i >= 1; --i)
                        {
                            if (slws.RowProperties.ContainsKey(i))
                            {
                                rp = slws.RowProperties[i];
                                if (rp.OutlineLevel > byCurrentOutlineLevel)
                                {
                                    rp.Hidden = false;
                                    slws.RowProperties[i] = rp.Clone();
                                }
                                else break;
                            }
                            else break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Indicates if the row has a thick top.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>True if the row has a thick top. False otherwise.</returns>
        public bool IsRowThickTopped(int RowIndex)
        {
            bool result = false;
            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                SLRowProperties rp = slws.RowProperties[RowIndex];
                result = rp.ThickTop;
            }

            return result;
        }

        /// <summary>
        /// Set the thick top property of the row.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="IsThickTopped">True if the row should have a thick top. False otherwise.</param>
        /// <returns>True if the row index is valid. False otherwise.</returns>
        public bool SetRowThickTop(int RowIndex, bool IsThickTopped)
        {
            return SetRowThickTop(RowIndex, RowIndex, IsThickTopped);
        }

        /// <summary>
        /// Set the thick top property of a range of rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <param name="IsThickTopped">True if the rows should have a thick top. False otherwise.</param>
        /// <returns>True if the row indices are valid. False otherwise.</returns>
        public bool SetRowThickTop(int StartRowIndex, int EndRowIndex, bool IsThickTopped)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            bool result = false;
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                result = true;
                int i = 0;
                SLRowProperties rp;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        rp.ThickTop = IsThickTopped;
                        slws.RowProperties[i] = rp;
                    }
                    else
                    {
                        rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                        rp.ThickTop = IsThickTopped;
                        slws.RowProperties.Add(i, rp);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Indicates if the row has a thick bottom.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>True if the row has a thick bottom. False otherwise.</returns>
        public bool IsRowThickBottomed(int RowIndex)
        {
            bool result = false;
            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                SLRowProperties rp = slws.RowProperties[RowIndex];
                result = rp.ThickBottom;
            }

            return result;
        }

        /// <summary>
        /// Set the thick bottom property of the row.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="IsThickBottomed">True if the row should have a thick bottom. False otherwise.</param>
        /// <returns>True if the row index is valid. False otherwise.</returns>
        public bool SetRowThickBottomed(int RowIndex, bool IsThickBottomed)
        {
            return SetRowThickBottomed(RowIndex, RowIndex, IsThickBottomed);
        }

        /// <summary>
        /// Set the thick bottom property of a range of rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <param name="IsThickBottomed">True if the rows should have a thick bottom. False otherwise.</param>
        /// <returns>True if the row indices are valid. False otherwise.</returns>
        public bool SetRowThickBottomed(int StartRowIndex, int EndRowIndex, bool IsThickBottomed)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            bool result = false;
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                result = true;
                int i = 0;
                SLRowProperties rp;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        rp.ThickBottom = IsThickBottomed;
                        slws.RowProperties[i] = rp;
                    }
                    else
                    {
                        rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                        rp.ThickBottom = IsThickBottomed;
                        slws.RowProperties.Add(i, rp);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Indicates if the row is showing phonetic information.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>True if the row is showing phonetic information. False otherwise.</returns>
        public bool IsRowShowingPhonetic(int RowIndex)
        {
            bool result = false;
            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                SLRowProperties rp = slws.RowProperties[RowIndex];
                result = rp.ShowPhonetic;
            }

            return result;
        }

        /// <summary>
        /// Set the show phonetic property for the row.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ShowPhonetic">True if the row should show phonetic information. False otherwise.</param>
        /// <returns>True if the row index is valid. False otherwise.</returns>
        public bool SetRowShowPhonetic(int RowIndex, bool ShowPhonetic)
        {
            return SetRowShowPhonetic(RowIndex, RowIndex, ShowPhonetic);
        }

        /// <summary>
        /// Set the show phonetic property for a range of rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <param name="ShowPhonetic">True if the rows should show phonetic information. False otherwise.</param>
        /// <returns>True if the row indices are valid. False otherwise.</returns>
        public bool SetRowShowPhonetic(int StartRowIndex, int EndRowIndex, bool ShowPhonetic)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            bool result = false;
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                result = true;
                int i = 0;
                SLRowProperties rp;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        rp.ShowPhonetic = ShowPhonetic;
                        slws.RowProperties[i] = rp;
                    }
                    else
                    {
                        rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                        rp.ShowPhonetic = ShowPhonetic;
                        slws.RowProperties.Add(i, rp);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Copy one row to another row.
        /// </summary>
        /// <param name="RowIndex">The row index of the row to be copied from.</param>
        /// <param name="AnchorRowIndex">The row index of the row to be copied to.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyRow(int RowIndex, int AnchorRowIndex)
        {
            return CopyRow(RowIndex, RowIndex, AnchorRowIndex, false);
        }

        /// <summary>
        /// Copy one row to another row.
        /// </summary>
        /// <param name="RowIndex">The row index of the row to be copied from.</param>
        /// <param name="AnchorRowIndex">The row index of the row to be copied to.</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyRow(int RowIndex, int AnchorRowIndex, bool ToCut)
        {
            return CopyRow(RowIndex, RowIndex, AnchorRowIndex, ToCut);
        }

        /// <summary>
        /// Copy a range of rows to another range, given the anchor row of the destination range (top row).
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row of the row range. This is typically the top row.</param>
        /// <param name="EndRowIndex">The row index of the end row of the row range. This is typically the bottom row.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor row.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyRow(int StartRowIndex, int EndRowIndex, int AnchorRowIndex)
        {
            return CopyRow(StartRowIndex, EndRowIndex, AnchorRowIndex, false);
        }

        /// <summary>
        /// Copy a range of rows to another range, given the anchor row of the destination range (top row).
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row of the row range. This is typically the top row.</param>
        /// <param name="EndRowIndex">The row index of the end row of the row range. This is typically the bottom row.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor row.</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyRow(int StartRowIndex, int EndRowIndex, int AnchorRowIndex, bool ToCut)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            bool result = false;
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit
                && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit
                && AnchorRowIndex >= 1 && AnchorRowIndex <= SLConstants.RowLimit
                && iStartRowIndex != AnchorRowIndex)
            {
                this.FlattenAllSharedCellFormula();

                result = true;

                int diff = AnchorRowIndex - iStartRowIndex;
                int i = 0;
                Dictionary<int, SLRowProperties> rows = new Dictionary<int, SLRowProperties>();
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rows[i + diff] = slws.RowProperties[i].Clone();
                        if (ToCut)
                        {
                            slws.RowProperties.Remove(i);
                        }
                    }
                }

                int AnchorEndRowIndex = AnchorRowIndex + iEndRowIndex - iStartRowIndex;
                // removing rows within destination "paste" operation
                List<int> rowkeys = slws.RowProperties.Keys.ToList<int>();
                foreach (int rkey in rowkeys)
                {
                    if (rkey >= AnchorRowIndex && rkey <= AnchorEndRowIndex)
                    {
                        slws.RowProperties.Remove(rkey);
                    }
                }

                foreach (var key in rows.Keys)
                {
                    slws.RowProperties[key] = rows[key].Clone();
                }

                SLCellWarehouse cells = new SLCellWarehouse();
                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                List<int> listColumnKeys;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        if (rowkey >= iStartRowIndex && rowkey <= iEndRowIndex)
                        {
                            cells.SetValue(rowkey + diff, colkey, slws.CellWarehouse.Cells[rowkey][colkey]);
                            if (ToCut)
                            {
                                slws.CellWarehouse.Remove(rowkey, colkey);
                            }
                        }
                    }
                }

                listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        // any cell within destination "paste" operation is taken out
                        if (rowkey >= AnchorRowIndex && rowkey <= AnchorEndRowIndex)
                        {
                            slws.CellWarehouse.Remove(rowkey, colkey);
                        }
                    }
                }

                int iNumberOfRows = iEndRowIndex - iStartRowIndex + 1;
                if (AnchorRowIndex <= iStartRowIndex) iNumberOfRows = -iNumberOfRows;

                SLCell c;
                listRowKeys = cells.Cells.Keys.ToList<int>();
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = cells.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        c = cells.Cells[rowkey][colkey];
                        //this.ProcessCellFormulaDelta(ref c, AnchorRowIndex, iNumberOfRows, -1, 0);
                        this.ProcessCellFormulaDelta(ref c, false, iStartRowIndex, 0, AnchorRowIndex, 0, false, false, false, 0, 0);
                        slws.CellWarehouse.SetValue(rowkey, colkey, c);
                    }
                }

                // TODO: tables!

                // cutting and pasting into a region with merged cells unmerges the existing merged cells
                // copying and pasting into a region with merged cells leaves existing merged cells alone.
                // Why does Excel do that? Don't know.
                // Will just standardise to leaving existing merged cells alone.
                List<SLMergeCell> mca = this.GetWorksheetMergeCells();
                foreach (SLMergeCell mc in mca)
                {
                    if (mc.StartRowIndex >= iStartRowIndex && mc.EndRowIndex <= iEndRowIndex)
                    {
                        if (ToCut)
                        {
                            slws.MergeCells.Remove(mc);
                        }
                        this.MergeWorksheetCells(mc.StartRowIndex + diff, mc.StartColumnIndex, mc.EndRowIndex + diff, mc.EndColumnIndex);
                    }
                }

                #region Calculation cells
                if (slwb.CalculationCells.Count > 0)
                {
                    // I don't know enough to fiddle with calculation chains. So I'm going to ignore it.
                    slwb.CalculationCells.Clear();

                    //List<int> listToDelete = new List<int>();
                    //int iRowLimit = AnchorRowIndex + iStartRowIndex - iEndRowIndex;
                    //for (i = 0; i < slwb.CalculationCells.Count; ++i)
                    //{
                    //    if (slwb.CalculationCells[i].SheetId == giSelectedWorksheetID)
                    //    {
                    //        if (ToCut && slwb.CalculationCells[i].RowIndex >= iStartRowIndex && slwb.CalculationCells[i].RowIndex <= iEndRowIndex)
                    //        {
                    //            // just remove because recalculation of cell references is too complicated...
                    //            if (!listToDelete.Contains(i)) listToDelete.Add(i);
                    //        }

                    //        if (slwb.CalculationCells[i].RowIndex >= AnchorRowIndex && slwb.CalculationCells[i].RowIndex <= iRowLimit)
                    //        {
                    //            // existing calculation cell lies within destination "paste" operation
                    //            if (!listToDelete.Contains(i)) listToDelete.Add(i);
                    //        }
                    //    }
                    //}

                    //for (i = listToDelete.Count - 1; i >= 0; --i)
                    //{
                    //    slwb.CalculationCells.RemoveAt(listToDelete[i]);
                    //}
                }
                #endregion

                // defined names is hard to calculate...
                // need to check the row and column indices based on the cell references within.
            }

            return result;
        }

        /// <summary>
        /// Insert one or more rows.
        /// </summary>
        /// <param name="StartRowIndex">Additional rows are inserted at this row index.</param>
        /// <param name="NumberOfRows">The number of rows to insert.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertRow(int StartRowIndex, int NumberOfRows)
        {
            if (NumberOfRows < 1) return false;

            bool result = false;
            if (StartRowIndex >= 1 && StartRowIndex <= SLConstants.RowLimit)
            {
                this.FlattenAllSharedCellFormula();

                result = true;
                int i = 0, iNewIndex = 0;

                int index = 0;
                int iRowIndex = -1;
                int iRowIndex2 = -1;

                #region Tables
                if (slws.Tables.Count > 0)
                {
                    foreach (SLTable t in slws.Tables)
                    {
                        iRowIndex = t.StartRowIndex;
                        iRowIndex2 = t.EndRowIndex;
                        this.AddRowColumnIndexDelta(StartRowIndex, NumberOfRows, true, ref iRowIndex, ref iRowIndex2);
                        if (iRowIndex != t.StartRowIndex || iRowIndex2 != t.EndRowIndex) t.IsNewTable = true;
                        t.StartRowIndex = iRowIndex;
                        t.EndRowIndex = iRowIndex2;

                        if (t.HasAutoFilter)
                        {
                            iRowIndex = t.AutoFilter.StartRowIndex;
                            iRowIndex2 = t.AutoFilter.EndRowIndex;
                            this.AddRowColumnIndexDelta(StartRowIndex, NumberOfRows, true, ref iRowIndex, ref iRowIndex2);
                            if (iRowIndex != t.AutoFilter.StartRowIndex || iRowIndex2 != t.AutoFilter.EndRowIndex) t.IsNewTable = true;
                            t.AutoFilter.StartRowIndex = iRowIndex;
                            t.AutoFilter.EndRowIndex = iRowIndex2;

                            if (t.AutoFilter.HasSortState)
                            {
                                iRowIndex = t.AutoFilter.SortState.StartRowIndex;
                                iRowIndex2 = t.AutoFilter.SortState.EndRowIndex;
                                this.AddRowColumnIndexDelta(StartRowIndex, NumberOfRows, true, ref iRowIndex, ref iRowIndex2);
                                if (iRowIndex != t.AutoFilter.SortState.StartRowIndex || iRowIndex2 != t.AutoFilter.SortState.EndRowIndex) t.IsNewTable = true;
                                t.AutoFilter.SortState.StartRowIndex = iRowIndex;
                                t.AutoFilter.SortState.EndRowIndex = iRowIndex2;
                            }
                        }

                        if (t.HasSortState)
                        {
                            iRowIndex = t.SortState.StartRowIndex;
                            iRowIndex2 = t.SortState.EndRowIndex;
                            this.AddRowColumnIndexDelta(StartRowIndex, NumberOfRows, true, ref iRowIndex, ref iRowIndex2);
                            if (iRowIndex != t.SortState.StartRowIndex || iRowIndex2 != t.SortState.EndRowIndex) t.IsNewTable = true;
                            t.SortState.StartRowIndex = iRowIndex;
                            t.SortState.EndRowIndex = iRowIndex2;
                        }
                    }
                }
                #endregion

                #region Row properties
                SLRowProperties rp;
                List<int> listRowIndex = slws.RowProperties.Keys.ToList<int>();
                // this sorting in descending order is crucial!
                // we move the data from after the insert range to their new reference keys
                // first, then we put in the new data, which will then have no data
                // key collision.
                listRowIndex.Sort();
                listRowIndex.Reverse();

                for (i = 0; i < listRowIndex.Count; ++i)
                {
                    index = listRowIndex[i];
                    if (index >= StartRowIndex)
                    {
                        rp = slws.RowProperties[index].Clone();
                        slws.RowProperties.Remove(index);
                        iNewIndex = index + NumberOfRows;
                        // if the new row is below the bottom limit of the worksheet,
                        // then it disappears into the ether...
                        if (iNewIndex <= SLConstants.RowLimit)
                        {
                            slws.RowProperties[iNewIndex] = rp.Clone();
                        }
                    }
                    else
                    {
                        // the rows before the start row are unaffected by the insertion.
                        // Because it's sorted in descending order, we can just break out.
                        break;
                    }
                }
                #endregion

                #region Cell data
                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                // 7 Jan 2017: need to sort descending. Similar logic to the RowProperties above.
                listRowKeys.Sort();
                listRowKeys.Reverse();
                List<int> listColumnKeys;

                SLCell c;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        c = slws.CellWarehouse.Cells[rowkey][colkey].Clone();
                        //this.ProcessCellFormulaDelta(ref c, StartRowIndex, NumberOfRows, -1, 0);
                        //this.ProcessCellFormulaDelta(ref c, false, StartRowIndex, 0, StartRowIndex + NumberOfRows, 0);
                        this.ProcessCellFormulaDelta(ref c, false, StartRowIndex, 0, StartRowIndex + NumberOfRows, 0, true, true, true, 0, 0);

                        if (rowkey >= StartRowIndex)
                        {
                            slws.CellWarehouse.Remove(rowkey, colkey);
                            iNewIndex = rowkey + NumberOfRows;
                            if (iNewIndex <= SLConstants.RowLimit)
                            {
                                slws.CellWarehouse.SetValue(iNewIndex, colkey, c);
                            }
                        }
                        else
                        {
                            slws.CellWarehouse.SetValue(rowkey, colkey, c);
                        }
                    }
                }

                #region Cell comments
                List<SLCellPoint> listCellRefKeys = slws.Comments.Keys.ToList<SLCellPoint>();
                // this sorting in descending order is crucial!
                listCellRefKeys.Sort(new SLCellReferencePointComparer());
                listCellRefKeys.Reverse();

                SLCellPoint pt;
                SLComment comm;
                for (i = 0; i < listCellRefKeys.Count; ++i)
                {
                    pt = listCellRefKeys[i];
                    comm = slws.Comments[pt];
                    if (pt.RowIndex >= StartRowIndex)
                    {
                        slws.Comments.Remove(pt);
                        iNewIndex = pt.RowIndex + NumberOfRows;
                        if (iNewIndex <= SLConstants.RowLimit)
                        {
                            slws.Comments[new SLCellPoint(iNewIndex, pt.ColumnIndex)] = comm;
                        }
                    }
                    // no else because there's nothing done
                }
                #endregion

                #endregion

                #region Merge cells
                if (slws.MergeCells.Count > 0)
                {
                    SLMergeCell mc;
                    for (i = 0; i < slws.MergeCells.Count; ++i)
                    {
                        mc = slws.MergeCells[i];
                        this.AddRowColumnIndexDelta(StartRowIndex, NumberOfRows, true, ref mc.iStartRowIndex, ref mc.iEndRowIndex);
                        slws.MergeCells[i] = mc;
                    }
                }
                #endregion

                #region Hyperlinks
                if (slws.Hyperlinks.Count > 0)
                {
                    SLHyperlink hl;
                    for (i = 0; i < slws.Hyperlinks.Count; ++i)
                    {
                        hl = slws.Hyperlinks[i];
                        iRowIndex = hl.Reference.StartRowIndex;
                        iRowIndex2 = hl.Reference.EndRowIndex;
                        this.AddRowColumnIndexDelta(StartRowIndex, NumberOfRows, true, ref iRowIndex, ref iRowIndex2);
                        hl.Reference = new SLCellPointRange(iRowIndex, hl.Reference.StartColumnIndex, iRowIndex2, hl.Reference.EndColumnIndex);
                        slws.Hyperlinks[i] = hl.Clone();
                    }
                }
                #endregion

                #region Drawings
                if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
                {
                    WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                    if (wsp.DrawingsPart != null)
                    {
                        bool bFound = false;
                        Xdr.TwoCellAnchor tcaNew;
                        Xdr.OneCellAnchor ocaNew;
                        Xdr.EditAsValues vEditAs = Xdr.EditAsValues.Absolute;
                        List<OpenXmlElement> listoxe = new List<OpenXmlElement>();
                        int iIndex = 0;

                        DrawingsPart dp = wsp.DrawingsPart;
                        foreach (OpenXmlElement oxe in dp.WorksheetDrawing.ChildElements)
                        {
                            if (oxe is Xdr.TwoCellAnchor)
                            {
                                tcaNew = (Xdr.TwoCellAnchor)oxe.CloneNode(true);

                                if (tcaNew.EditAs == null) vEditAs = Xdr.EditAsValues.TwoCell;
                                else vEditAs = tcaNew.EditAs.Value;

                                if (vEditAs == Xdr.EditAsValues.TwoCell)
                                {
                                    if (tcaNew.FromMarker != null && tcaNew.FromMarker.RowId != null)
                                    {
                                        iIndex = Convert.ToInt32(tcaNew.FromMarker.RowId.Text);
                                        // the index is 0-based while the spreadsheet index is 1-based.
                                        if ((iIndex + 1) >= StartRowIndex)
                                        {
                                            iIndex += NumberOfRows;
                                            bFound = true;
                                            tcaNew.FromMarker.RowId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                        }
                                    }

                                    if (tcaNew.ToMarker != null && tcaNew.ToMarker.RowId != null)
                                    {
                                        iIndex = Convert.ToInt32(tcaNew.ToMarker.RowId.Text);
                                        // the index is 0-based while the spreadsheet index is 1-based.
                                        if ((iIndex + 1) >= StartRowIndex)
                                        {
                                            iIndex += NumberOfRows;
                                            bFound = true;
                                            tcaNew.ToMarker.RowId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                        }
                                    }
                                }
                                else if (vEditAs == Xdr.EditAsValues.OneCell)
                                {
                                    if (tcaNew.FromMarker != null && tcaNew.FromMarker.RowId != null)
                                    {
                                        iIndex = Convert.ToInt32(tcaNew.FromMarker.RowId.Text);
                                        // the index is 0-based while the spreadsheet index is 1-based.
                                        if ((iIndex + 1) >= StartRowIndex)
                                        {
                                            iIndex += NumberOfRows;
                                            bFound = true;
                                            tcaNew.FromMarker.RowId.Text = iIndex.ToString(CultureInfo.InvariantCulture);

                                            // if the from marker is moved, then the to marker has to move too
                                            if (tcaNew.ToMarker != null && tcaNew.ToMarker.RowId != null)
                                            {
                                                iIndex = Convert.ToInt32(tcaNew.ToMarker.RowId.Text);
                                                iIndex += NumberOfRows;
                                                tcaNew.ToMarker.RowId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                            }
                                        }
                                    }
                                }

                                // Need to do for Transform for the child elements?

                                listoxe.Add(tcaNew.CloneNode(true));
                            }
                            else if (oxe is Xdr.OneCellAnchor)
                            {
                                ocaNew = (Xdr.OneCellAnchor)oxe.CloneNode(true);
                                if (ocaNew.FromMarker != null && ocaNew.FromMarker.RowId != null)
                                {
                                    iIndex = Convert.ToInt32(ocaNew.FromMarker.RowId.Text);
                                    // the index is 0-based while the spreadsheet index is 1-based.
                                    if ((iIndex + 1) >= StartRowIndex)
                                    {
                                        iIndex += NumberOfRows;
                                        bFound = true;
                                        ocaNew.FromMarker.RowId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                    }
                                }

                                // Need to do for Transform for the child elements?

                                listoxe.Add(ocaNew.CloneNode(true));
                            }
                            else
                            {
                                listoxe.Add(oxe.CloneNode(true));
                            }
                        }

                        if (bFound)
                        {
                            wsp.DrawingsPart.WorksheetDrawing.RemoveAllChildren();
                            foreach (OpenXmlElement oxe in listoxe)
                            {
                                wsp.DrawingsPart.WorksheetDrawing.Append(oxe.CloneNode(true));
                            }
                            wsp.DrawingsPart.WorksheetDrawing.Save();
                        }
                    }
                }
                #endregion

                // TODO: chart series references

                #region Calculation chain
                if (slwb.CalculationCells.Count > 0)
                {
                    // I don't know enough to fiddle with calculation chains. So I'm going to ignore it.
                    slwb.CalculationCells.Clear();

                    //foreach (SLCalculationCell cc in slwb.CalculationCells)
                    //{
                    //    if (cc.SheetId == giSelectedWorksheetID)
                    //    {
                    //        iRowIndex = cc.RowIndex;
                    //        // don't need this but assign something anyway...
                    //        iRowIndex2 = SLConstants.RowLimit;

                    //        this.AddRowColumnIndexDelta(StartRowIndex, NumberOfRows, true, ref iRowIndex, ref iRowIndex2);
                    //        cc.RowIndex = iRowIndex;
                    //    }
                    //}
                }
                #endregion

                #region Defined names
                if (slwb.DefinedNames.Count > 0)
                {
                    //string sDefinedNameText = string.Empty;
                    //foreach (SLDefinedName d in slwb.DefinedNames)
                    //{
                    //    sDefinedNameText = d.Text;
                    //    sDefinedNameText = AddDeleteCellFormulaDelta(sDefinedNameText, StartRowIndex, NumberOfRows, -1, 0);
                    //    sDefinedNameText = AddDeleteDefinedNameRowColumnRangeDelta(sDefinedNameText, true, StartRowIndex, NumberOfRows);
                    //    d.Text = sDefinedNameText;
                    //}

                    // not doing anything with this but using it to pass to the function
                    bool bHasError = false;
                    foreach (SLDefinedName d in slwb.DefinedNames)
                    {
                        d.Text = AdjustCellFormulaDelta(d.Text, false, StartRowIndex, 0, StartRowIndex + NumberOfRows, 0, true, false, false, false, 0, 0, out bHasError);
                    }
                }
                #endregion

                #region Sparklines
                if (slws.SparklineGroups.Count > 0)
                {
                    SLSparkline spk;
                    foreach (SLSparklineGroup spkgrp in slws.SparklineGroups)
                    {
                        if (spkgrp.DateAxis && spkgrp.DateWorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            this.AddRowColumnIndexDelta(StartRowIndex, NumberOfRows, true, ref spkgrp.DateStartRowIndex, ref spkgrp.DateEndRowIndex);
                        }

                        // starting from the end is important because we might be deleting!
                        for (i = spkgrp.Sparklines.Count - 1; i >= 0; --i)
                        {
                            spk = spkgrp.Sparklines[i];

                            if (spk.LocationRowIndex >= StartRowIndex)
                            {
                                iNewIndex = spk.LocationRowIndex + NumberOfRows;
                                if (iNewIndex <= SLConstants.RowLimit)
                                {
                                    spk.LocationRowIndex = iNewIndex;
                                }
                                else
                                {
                                    // out of range!
                                    spkgrp.Sparklines.RemoveAt(i);
                                    continue;
                                }
                            }
                            // else the location is before the start row so don't have to do anything

                            // process only if the data source is on the currently selected worksheet
                            if (spk.WorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                this.AddRowColumnIndexDelta(StartRowIndex, NumberOfRows, true, ref spk.StartRowIndex, ref spk.EndRowIndex);
                            }

                            spkgrp.Sparklines[i] = spk;
                        }
                    }
                }
                #endregion
            }

            return result;
        }

        /// <summary>
        /// Delete one or more rows.
        /// </summary>
        /// <param name="StartRowIndex">Rows will be deleted from this row index, including this row itself.</param>
        /// <param name="NumberOfRows">Number of rows to delete.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool DeleteRow(int StartRowIndex, int NumberOfRows)
        {
            if (NumberOfRows < 1) return false;

            bool result = false;
            if (StartRowIndex >= 1 && StartRowIndex <= SLConstants.RowLimit)
            {
                this.FlattenAllSharedCellFormula();

                result = true;
                int i = 0, iNewIndex = 0;
                int iEndRowIndex = StartRowIndex + NumberOfRows - 1;
                if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
                // this autocorrects in the case of overshooting the row limit
                int iNumberOfRows = iEndRowIndex - StartRowIndex + 1;

                WorksheetPart wsp;

                int index = 0;
                int iRowIndex = -1;
                int iRowIndex2 = -1;

                // tables part has to be at the beginning because we need to check if
                // the header row of any table is within delete range, BUT the whole of that
                // table is NOT within delete range (meaning we're deleting the header row
                // without deleting the whole table), and we exit the function (because
                // Excel doesn't allow this behaviour).
                #region Tables
                if (slws.Tables.Count > 0)
                {
                    SLTable t;
                    #region Table header check
                    for (i = 0; i < slws.Tables.Count; ++i)
                    {
                        t = slws.Tables[i];
                        if (t.HeaderRowCount > 0)
                        {
                            if (StartRowIndex <= t.StartRowIndex && t.StartRowIndex <= iEndRowIndex && iEndRowIndex < t.EndRowIndex)
                            {
                                // the delete range includes a header row, BUT does not
                                // delete the whole table.
                                return false;
                            }
                        }

                        // check if the delete range contains the body of the table
                        // Excel allows this, but keeps an empty row afterwards.
                        // This means even though 6 rows are to be deleted, Excel only deletes
                        // 5 rows, leaving an empty row after that.
                        // Without visual feedback, this is difficult to keep track from the calling
                        // program, so we'll just disallow this.
                        if (t.HeaderRowCount > 0 && t.TotalsRowCount > 0)
                        {
                            if ((StartRowIndex == (t.StartRowIndex + 1)) && (iEndRowIndex == (t.EndRowIndex - 1)))
                            {
                                return false;
                            }
                        }
                        else if (t.HeaderRowCount > 0 && t.TotalsRowCount == 0)
                        {
                            if ((StartRowIndex == (t.StartRowIndex + 1)) && (iEndRowIndex >= t.EndRowIndex))
                            {
                                return false;
                            }
                        }
                        else if (t.HeaderRowCount == 0 && t.TotalsRowCount > 0)
                        {
                            if ((StartRowIndex <= t.StartRowIndex) && (iEndRowIndex == (t.EndRowIndex - 1)))
                            {
                                return false;
                            }
                        }

                        // else there are no header rows or totals row.
                        // and if the body of the table is within delete range,
                        // then it's taken care of below.
                    }
                    #endregion

                    List<int> listTablesToDelete = new List<int>();
                    for (i = 0; i < slws.Tables.Count; ++i)
                    {
                        t = slws.Tables[i];
                        if (StartRowIndex <= t.StartRowIndex && t.EndRowIndex <= iEndRowIndex)
                        {
                            // table is completely within delete range, so delete the whole table
                            listTablesToDelete.Add(i);
                            continue;
                        }
                        else
                        {
                            if (t.TotalsRowCount > 0)
                            {
                                // the totals row is within delete range
                                if (StartRowIndex <= t.EndRowIndex && t.EndRowIndex <= iEndRowIndex)
                                {
                                    // should be just minus 1, but we'll just do this instead...
                                    t.EndRowIndex -= (int)t.TotalsRowCount;
                                    t.TotalsRowCount = 0;
                                    t.IsNewTable = true;
                                }
                            }

                            iRowIndex = t.StartRowIndex;
                            iRowIndex2 = t.EndRowIndex;
                            this.DeleteRowColumnIndexDelta(StartRowIndex, iEndRowIndex, iNumberOfRows, ref iRowIndex, ref iRowIndex2);
                            if (iRowIndex != t.StartRowIndex || iRowIndex2 != t.EndRowIndex) t.IsNewTable = true;
                            t.StartRowIndex = iRowIndex;
                            t.EndRowIndex = iRowIndex2;
                        }

                        if (t.HasAutoFilter)
                        {
                            // if the autofilter range is completely within delete range,
                            // then it's already taken care off above.
                            iRowIndex = t.AutoFilter.StartRowIndex;
                            iRowIndex2 = t.AutoFilter.EndRowIndex;
                            this.DeleteRowColumnIndexDelta(StartRowIndex, iEndRowIndex, iNumberOfRows, ref iRowIndex, ref iRowIndex2);
                            if (iRowIndex != t.AutoFilter.StartRowIndex || iRowIndex2 != t.AutoFilter.EndRowIndex) t.IsNewTable = true;
                            t.AutoFilter.StartRowIndex = iRowIndex;
                            t.AutoFilter.EndRowIndex = iRowIndex2;

                            if (t.AutoFilter.HasSortState)
                            {
                                // if the sort state range is completely within delete range,
                                // then it's already taken care off above.
                                iRowIndex = t.AutoFilter.SortState.StartRowIndex;
                                iRowIndex2 = t.AutoFilter.SortState.EndRowIndex;
                                this.DeleteRowColumnIndexDelta(StartRowIndex, iEndRowIndex, iNumberOfRows, ref iRowIndex, ref iRowIndex2);
                                if (iRowIndex != t.AutoFilter.SortState.StartRowIndex || iRowIndex2 != t.AutoFilter.SortState.EndRowIndex) t.IsNewTable = true;
                                t.AutoFilter.SortState.StartRowIndex = iRowIndex;
                                t.AutoFilter.SortState.EndRowIndex = iRowIndex2;
                            }
                        }

                        if (t.HasSortState)
                        {
                            // if the sort state range is completely within delete range,
                            // then it's already taken care off above.
                            iRowIndex = t.SortState.StartRowIndex;
                            iRowIndex2 = t.SortState.EndRowIndex;
                            this.DeleteRowColumnIndexDelta(StartRowIndex, iEndRowIndex, iNumberOfRows, ref iRowIndex, ref iRowIndex2);
                            if (iRowIndex != t.SortState.StartRowIndex || iRowIndex2 != t.SortState.EndRowIndex) t.IsNewTable = true;
                            t.SortState.StartRowIndex = iRowIndex;
                            t.SortState.EndRowIndex = iRowIndex2;
                        }
                    }

                    if (listTablesToDelete.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
                        {
                            wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                            string sTableRelID = string.Empty;
                            string sTableName = string.Empty;
                            uint iTableID = 0;
                            for (i = listTablesToDelete.Count - 1; i >= 0; --i)
                            {
                                // remove IDs and table names from the spreadsheet unique lists
                                iTableID = slws.Tables[listTablesToDelete[i]].Id;
                                if (slwb.TableIds.Contains(iTableID)) slwb.TableIds.Remove(iTableID);

                                sTableName = slws.Tables[listTablesToDelete[i]].DisplayName;
                                if (slwb.TableNames.Contains(sTableName)) slwb.TableNames.Remove(sTableName);

                                sTableRelID = slws.Tables[listTablesToDelete[i]].RelationshipID;
                                if (sTableRelID.Length > 0)
                                {
                                    wsp.DeletePart(sTableRelID);
                                }
                                slws.Tables.RemoveAt(listTablesToDelete[i]);
                            }
                        }
                    }
                }
                #endregion

                #region Row properties
                SLRowProperties rp;
                List<int> listRowIndex = slws.RowProperties.Keys.ToList<int>();
                // this sorting in ascending order is crucial!
                // we're removing data within delete range, and if the data is after the delete range,
                // then the key is modified. Since the data within delete range is removed,
                // there won't be data key collisions.
                listRowIndex.Sort();

                for (i = 0; i < listRowIndex.Count; ++i)
                {
                    index = listRowIndex[i];
                    if (index >= StartRowIndex && index <= iEndRowIndex)
                    {
                        slws.RowProperties.Remove(index);
                    }
                    else if (index > iEndRowIndex)
                    {
                        rp = slws.RowProperties[index];
                        slws.RowProperties.Remove(index);
                        iNewIndex = index - iNumberOfRows;
                        slws.RowProperties[iNewIndex] = rp.Clone();
                    }

                    // the rows before the start row are unaffected by the deleting.
                }
                #endregion

                #region Cell data
                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                // 2 April 2017: Need to sort or the row keys are "random" (depending on
                // how the OS assign memory and order and such)
                // This made the rows "random"
                listRowKeys.Sort();
                List<int> listColumnKeys;

                SLCell c;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    listColumnKeys.Sort();
                    foreach (int colkey in listColumnKeys)
                    {
                        c = slws.CellWarehouse.Cells[rowkey][colkey].Clone();

                        if (StartRowIndex <= rowkey && rowkey <= iEndRowIndex)
                        {
                            slws.CellWarehouse.Remove(rowkey, colkey);
                        }
                        else if (rowkey > iEndRowIndex)
                        {
                            slws.CellWarehouse.Remove(rowkey, colkey);
                            iNewIndex = rowkey - iNumberOfRows;
                            this.ProcessCellFormulaDelta(ref c, false, StartRowIndex, 0, StartRowIndex - iNumberOfRows, 0, true, false, true, StartRowIndex, iEndRowIndex);
                            slws.CellWarehouse.SetValue(iNewIndex, colkey, c);
                        }
                        //else
                        //{
                        //    slws.CellWarehouse.SetValue(rowkey, colkey, c);
                        //}
                    }
                }

                #region Cell comments
                List<SLCellPoint> listCellRefKeys = slws.Comments.Keys.ToList<SLCellPoint>();
                // this sorting in ascending order is crucial!
                listCellRefKeys.Sort(new SLCellReferencePointComparer());

                SLCellPoint pt;
                SLComment comm;
                for (i = 0; i < listCellRefKeys.Count; ++i)
                {
                    pt = listCellRefKeys[i];
                    comm = slws.Comments[pt];
                    if (StartRowIndex <= pt.RowIndex && pt.RowIndex <= iEndRowIndex)
                    {
                        slws.Comments.Remove(pt);
                    }
                    else if (pt.RowIndex > iEndRowIndex)
                    {
                        slws.Comments.Remove(pt);
                        iNewIndex = pt.RowIndex - iNumberOfRows;
                        slws.Comments[new SLCellPoint(iNewIndex, pt.ColumnIndex)] = comm;
                    }
                    // no else because there's nothing done
                }
                #endregion

                #endregion

                #region Merge cells
                if (slws.MergeCells.Count > 0)
                {
                    SLMergeCell mc;
                    // starting from the end is crucial because we might be removing items
                    for (i = slws.MergeCells.Count - 1; i >= 0; --i)
                    {
                        mc = slws.MergeCells[i];
                        if (mc.iStartRowIndex >= StartRowIndex && mc.iEndRowIndex <= iEndRowIndex)
                        {
                            // merge cell is completely within delete range
                            slws.MergeCells.RemoveAt(i);
                        }
                        else
                        {
                            this.DeleteRowColumnIndexDelta(StartRowIndex, iEndRowIndex, iNumberOfRows, ref mc.iStartRowIndex, ref mc.iEndRowIndex);
                            slws.MergeCells[i] = mc;
                        }
                    }
                }
                #endregion

                #region Hyperlinks
                if (slws.Hyperlinks.Count > 0)
                {
                    List<SLCellPoint> deletepoints = new List<SLCellPoint>();
                    SLHyperlink hl;
                    for (i = slws.Hyperlinks.Count - 1; i >= 0; --i)
                    {
                        hl = slws.Hyperlinks[i];
                        if (hl.Reference.StartRowIndex >= StartRowIndex && hl.Reference.EndRowIndex <= iEndRowIndex)
                        {
                            // hyperlink is completely within delete range
                            // We use a list to take of the points because we also need to remove any existing
                            // hyperlink relationship. It's easier to just use the remove hyperlink function.
                            deletepoints.Add(new SLCellPoint(hl.Reference.StartRowIndex, hl.Reference.StartColumnIndex));
                        }
                        else
                        {
                            iRowIndex = hl.Reference.StartRowIndex;
                            iRowIndex2 = hl.Reference.EndRowIndex;
                            this.DeleteRowColumnIndexDelta(StartRowIndex, iEndRowIndex, iNumberOfRows, ref iRowIndex, ref iRowIndex2);
                            hl.Reference = new SLCellPointRange(iRowIndex, hl.Reference.StartColumnIndex, iRowIndex2, hl.Reference.EndColumnIndex);
                            slws.Hyperlinks[i] = hl.Clone();
                        }
                    }

                    foreach (SLCellPoint hyperlinkpt in deletepoints)
                    {
                        this.RemoveHyperlink(hyperlinkpt.RowIndex, hyperlinkpt.ColumnIndex);
                    }
                }
                #endregion

                #region Drawings
                if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
                {
                    wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                    if (wsp.DrawingsPart != null)
                    {
                        bool bFound = false;
                        bool bToFindCumulative = false;
                        Xdr.TwoCellAnchor tcaNew;
                        Xdr.OneCellAnchor ocaNew;
                        Xdr.EditAsValues vEditAs = Xdr.EditAsValues.Absolute;
                        List<OpenXmlElement> listoxe = new List<OpenXmlElement>();
                        int iIndex = 0;
                        int iIndex2 = 0;

                        long lDefaultHeight = slws.SheetFormatProperties.DefaultRowHeightInEMU;
                        Dictionary<int, long> dictRowHeight = new Dictionary<int, long>();
                        List<int> listindex = slws.RowProperties.Keys.ToList<int>();
                        foreach (int rowindex in listindex)
                        {
                            rp = slws.RowProperties[rowindex];
                            dictRowHeight[rowindex] = rp.HeightInEMU;
                        }

                        Dictionary<int, long> dictCumulativeLength = new Dictionary<int, long>();

                        long lStartDelete = 0, lEndDelete = 0;
                        long lAnchorStartDelete = 0, lAnchorEndDelete = 0;

                        long lCumulative = 0;
                        for (iIndex = 1; iIndex <= iEndRowIndex; ++iIndex)
                        {
                            if (dictRowHeight.ContainsKey(iIndex))
                            {
                                lCumulative += dictRowHeight[iIndex];
                            }
                            else
                            {
                                lCumulative += lDefaultHeight;
                            }

                            dictCumulativeLength[iIndex] = lCumulative;
                        }

                        // we're getting the start and end "lines" of the delete range in terms of EMUs.
                        // The start line is at the top of the delete range, which is the bottom of the
                        // previous row. The end line is at the bottom of the delete range, which
                        // is exactly the current row.

                        if (StartRowIndex > 1) lStartDelete = dictCumulativeLength[StartRowIndex - 1];
                        else lStartDelete = 0;

                        lEndDelete = dictCumulativeLength[iEndRowIndex];

                        DrawingsPart dp = wsp.DrawingsPart;
                        foreach (OpenXmlElement oxe in dp.WorksheetDrawing.ChildElements)
                        {
                            if (oxe is Xdr.TwoCellAnchor)
                            {
                                tcaNew = (Xdr.TwoCellAnchor)oxe.CloneNode(true);

                                if (tcaNew.EditAs == null) vEditAs = Xdr.EditAsValues.TwoCell;
                                else vEditAs = tcaNew.EditAs.Value;

                                if (vEditAs != Xdr.EditAsValues.Absolute
                                    && tcaNew.FromMarker != null && tcaNew.FromMarker.RowId != null
                                    && tcaNew.ToMarker != null && tcaNew.ToMarker.RowId != null)
                                {
                                    iIndex = Convert.ToInt32(tcaNew.FromMarker.RowId.Text);
                                    iIndex2 = Convert.ToInt32(tcaNew.ToMarker.RowId.Text);

                                    if (iIndex2 + 1 > dictCumulativeLength.Count)
                                    {
                                        lCumulative = dictCumulativeLength[dictCumulativeLength.Count];
                                        for (i = dictCumulativeLength.Count + 1; i <= iIndex2 + 1; ++i)
                                        {
                                            if (dictRowHeight.ContainsKey(i))
                                            {
                                                lCumulative += dictRowHeight[i];
                                            }
                                            else
                                            {
                                                lCumulative += lDefaultHeight;
                                            }
                                            dictCumulativeLength[i] = lCumulative;
                                        }
                                    }

                                    // the index is 0-based while the spreadsheet index is 1-based.

                                    if (iIndex2 > 0) lAnchorEndDelete = dictCumulativeLength[iIndex2];
                                    else lAnchorEndDelete = 0;

                                    if (tcaNew.ToMarker.RowOffset != null)
                                    {
                                        lAnchorEndDelete += Convert.ToInt64(tcaNew.ToMarker.RowOffset.Text);
                                    }

                                    if (lAnchorEndDelete > dictCumulativeLength[dictCumulativeLength.Count])
                                    {
                                        lCumulative = dictCumulativeLength[dictCumulativeLength.Count];
                                        while (lAnchorEndDelete > lCumulative)
                                        {
                                            i = dictCumulativeLength.Count + 1;
                                            if (dictRowHeight.ContainsKey(i))
                                            {
                                                lCumulative += dictRowHeight[i];
                                            }
                                            else
                                            {
                                                lCumulative += lDefaultHeight;
                                            }
                                            dictCumulativeLength[i] = lCumulative;
                                        }
                                    }

                                    // assume the ToMarker index is after the FromMarker, and that after
                                    // taking into account the offset, it's still before the ToMarker,
                                    // so we don't do any more checks

                                    if (iIndex > 0) lAnchorStartDelete = dictCumulativeLength[iIndex];
                                    else lAnchorStartDelete = 0;

                                    if (tcaNew.FromMarker.RowOffset != null)
                                    {
                                        lAnchorStartDelete += Convert.ToInt64(tcaNew.FromMarker.RowOffset.Text);
                                    }

                                    bToFindCumulative = false;
                                    if (vEditAs == Xdr.EditAsValues.TwoCell)
                                    {
                                        if (lStartDelete <= lAnchorStartDelete && lAnchorEndDelete <= lEndDelete)
                                        {
                                            // completely within delete range
                                            // Not sure how to handle any hyperlinks tied to drawing or physical media files.
                                            // So will just make the drawing very very small. Say 1 EMU.
                                            // That's practically invisible.
                                            tcaNew.FromMarker.RowOffset = new Xdr.RowOffset("0");
                                            tcaNew.ToMarker.RowId = new Xdr.RowId(tcaNew.FromMarker.RowId.Text);
                                            tcaNew.ToMarker.RowOffset = new Xdr.RowOffset("1");
                                            bFound = true;
                                            bToFindCumulative = false;
                                        }
                                        else if (lEndDelete <= lAnchorStartDelete)
                                        {
                                            // delete range is before drawing
                                            lCumulative = lEndDelete - lStartDelete;
                                            lAnchorEndDelete -= lCumulative;
                                            lAnchorStartDelete -= lCumulative;

                                            bToFindCumulative = true;
                                        }
                                        else if (lStartDelete <= lAnchorStartDelete && lAnchorStartDelete <= lEndDelete && lEndDelete <= lAnchorEndDelete)
                                        {
                                            // top part is within delete range
                                            lCumulative = lEndDelete - lAnchorStartDelete;
                                            lAnchorEndDelete -= lCumulative;

                                            // put the drawing flush at the start of delete range
                                            lCumulative = lAnchorEndDelete - lAnchorStartDelete;
                                            lAnchorStartDelete = lStartDelete;
                                            lAnchorEndDelete = lAnchorStartDelete + lCumulative;

                                            bToFindCumulative = true;
                                        }
                                        else if (lAnchorStartDelete <= lStartDelete && lEndDelete <= lAnchorEndDelete)
                                        {
                                            // delete range is within drawing
                                            lCumulative = lEndDelete - lStartDelete;
                                            lAnchorEndDelete -= lCumulative;

                                            bToFindCumulative = true;
                                        }
                                        else if (lAnchorStartDelete <= lStartDelete && lStartDelete <= lAnchorEndDelete && lAnchorEndDelete <= lEndDelete)
                                        {
                                            // bottom part is within delete range
                                            lAnchorEndDelete = lStartDelete;

                                            bToFindCumulative = true;
                                        }
                                        // else the delete range is beyond the drawing, so nothing needs to be done
                                    }
                                    else if (vEditAs == Xdr.EditAsValues.OneCell)
                                    {
                                        if (lStartDelete <= lAnchorStartDelete)
                                        {
                                            // hold temporarily
                                            lCumulative = lAnchorStartDelete;

                                            lAnchorStartDelete -= (lEndDelete - lStartDelete);
                                            if (lAnchorStartDelete < lStartDelete) lAnchorStartDelete = lStartDelete;

                                            lAnchorEndDelete -= (lCumulative - lAnchorStartDelete);
                                            // the above algorithm makes it such that we don't overshoot the start delete
                                            // range. Normally, we'd only have the anchored start (FromMarker), but
                                            // TwoCellAnchor's have the ToMarker too, which is a pain in the Adam's apple.

                                            bToFindCumulative = true;
                                        }
                                    }

                                    if (bToFindCumulative)
                                    {
                                        iIndex = -1;
                                        iIndex2 = -1;
                                        for (i = dictCumulativeLength.Count; i >= 0; --i)
                                        {
                                            if (i == 0)
                                            {
                                                if (iIndex < 0)
                                                {
                                                    iIndex = 0;
                                                    tcaNew.FromMarker.RowId = new Xdr.RowId(iIndex.ToString(CultureInfo.InvariantCulture));
                                                    tcaNew.FromMarker.RowOffset = new Xdr.RowOffset(lAnchorStartDelete.ToString(CultureInfo.InvariantCulture));
                                                }

                                                if (iIndex2 < 0)
                                                {
                                                    iIndex2 = 0;
                                                    tcaNew.ToMarker.RowId = new Xdr.RowId(iIndex2.ToString(CultureInfo.InvariantCulture));
                                                    tcaNew.ToMarker.RowOffset = new Xdr.RowOffset(lAnchorEndDelete.ToString(CultureInfo.InvariantCulture));
                                                }
                                            }

                                            if (iIndex < 0 && lAnchorStartDelete >= dictCumulativeLength[i])
                                            {
                                                iIndex = i;
                                                tcaNew.FromMarker.RowId = new Xdr.RowId(iIndex.ToString(CultureInfo.InvariantCulture));
                                                tcaNew.FromMarker.RowOffset = new Xdr.RowOffset((lAnchorStartDelete - dictCumulativeLength[i]).ToString(CultureInfo.InvariantCulture));
                                            }
                                            if (iIndex2 < 0 && lAnchorEndDelete >= dictCumulativeLength[i])
                                            {
                                                iIndex2 = i;
                                                tcaNew.ToMarker.RowId = new Xdr.RowId(iIndex2.ToString(CultureInfo.InvariantCulture));
                                                tcaNew.ToMarker.RowOffset = new Xdr.RowOffset((lAnchorEndDelete - dictCumulativeLength[i]).ToString(CultureInfo.InvariantCulture));
                                            }

                                            if (iIndex >= 0 && iIndex2 >= 0) break;
                                        }

                                        if (iIndex >= 0 && iIndex2 >= 0)
                                        {
                                            bFound = true;
                                        }
                                    }
                                }

                                // Need to do for Transform for the child elements?

                                listoxe.Add(tcaNew.CloneNode(true));
                            }
                            else if (oxe is Xdr.OneCellAnchor)
                            {
                                ocaNew = (Xdr.OneCellAnchor)oxe.CloneNode(true);
                                if (ocaNew.FromMarker != null && ocaNew.FromMarker.RowId != null)
                                {
                                    iIndex = Convert.ToInt32(ocaNew.FromMarker.RowId.Text);
                                    // the index is 0-based while the spreadsheet index is 1-based.
                                    if ((iIndex + 1) >= StartRowIndex)
                                    {
                                        iIndex -= iNumberOfRows;
                                        // Excel seems to stop drawings from overshooting
                                        if ((iIndex + 1) < StartRowIndex)
                                        {
                                            iIndex = StartRowIndex - 1;
                                            // to put drawing flush with row
                                            if (ocaNew.Extent != null && ocaNew.Extent.Cy != null)
                                            {
                                                ocaNew.Extent.Cy = 0;
                                            }
                                        }
                                        bFound = true;
                                        ocaNew.FromMarker.RowId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                    }
                                }

                                // Need to do for Transform for the child elements?

                                listoxe.Add(ocaNew.CloneNode(true));
                            }
                            else
                            {
                                listoxe.Add(oxe.CloneNode(true));
                            }
                        }

                        bFound = true;
                        if (bFound)
                        {
                            wsp.DrawingsPart.WorksheetDrawing.RemoveAllChildren();
                            foreach (OpenXmlElement oxe in listoxe)
                            {
                                wsp.DrawingsPart.WorksheetDrawing.Append(oxe.CloneNode(true));
                            }
                            wsp.DrawingsPart.WorksheetDrawing.Save();
                        }
                    }
                }
                #endregion

                // TODO: chart series references

                #region Calculation chain
                if (slwb.CalculationCells.Count > 0)
                {
                    // I don't know enough to fiddle with calculation chains. So I'm going to ignore it.
                    slwb.CalculationCells.Clear();

                    //List<int> listToDelete = new List<int>();
                    //for (i = 0; i < slwb.CalculationCells.Count; ++i)
                    //{
                    //    if (slwb.CalculationCells[i].SheetId == giSelectedWorksheetID)
                    //    {
                    //        if (StartRowIndex <= slwb.CalculationCells[i].RowIndex && slwb.CalculationCells[i].RowIndex <= iEndRowIndex)
                    //        {
                    //            listToDelete.Add(i);
                    //        }
                    //        else if (iEndRowIndex < slwb.CalculationCells[i].RowIndex)
                    //        {
                    //            slwb.CalculationCells[i].RowIndex -= iNumberOfRows;
                    //        }
                    //    }
                    //}

                    //// start from the back because we're deleting elements and we don't want
                    //// the indices to get messed up.
                    //for (i = listToDelete.Count - 1; i >= 0; --i)
                    //{
                    //    slwb.CalculationCells.RemoveAt(listToDelete[i]);
                    //}
                }
                #endregion

                #region Defined names
                if (slwb.DefinedNames.Count > 0)
                {
                    //string sDefinedNameText = string.Empty;
                    //foreach (SLDefinedName d in slwb.DefinedNames)
                    //{
                    //    sDefinedNameText = d.Text;
                    //    sDefinedNameText = AddDeleteCellFormulaDelta(sDefinedNameText, StartRowIndex, -NumberOfRows, -1, 0);
                    //    sDefinedNameText = AddDeleteDefinedNameRowColumnRangeDelta(sDefinedNameText, true, StartRowIndex, -NumberOfRows);
                    //    d.Text = sDefinedNameText;
                    //}

                    // not doing anything with this but using it to pass to the function
                    bool bHasError = false;
                    foreach (SLDefinedName d in slwb.DefinedNames)
                    {
                        d.Text = AdjustCellFormulaDelta(d.Text, false, StartRowIndex, 0, StartRowIndex - NumberOfRows, 0, true, true, false, true, StartRowIndex, iEndRowIndex, out bHasError);
                    }
                }
                #endregion

                #region Sparklines
                if (slws.SparklineGroups.Count > 0)
                {
                    SLSparkline spk;
                    foreach (SLSparklineGroup spkgrp in slws.SparklineGroups)
                    {
                        if (spkgrp.DateAxis && spkgrp.DateWorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (StartRowIndex <= spkgrp.DateStartRowIndex && spkgrp.DateEndRowIndex <= iEndRowIndex)
                            {
                                // the whole date range is completely within delete range
                                spkgrp.DateAxis = false;
                            }
                            else
                            {
                                this.DeleteRowColumnIndexDelta(StartRowIndex, iEndRowIndex, iNumberOfRows, ref spkgrp.DateStartRowIndex, ref spkgrp.DateEndRowIndex);
                            }
                        }

                        // starting from the end is important because we might be deleting!
                        for (i = spkgrp.Sparklines.Count - 1; i >= 0; --i)
                        {
                            spk = spkgrp.Sparklines[i];

                            if (StartRowIndex <= spk.LocationRowIndex && spk.LocationRowIndex <= iEndRowIndex)
                            {
                                spkgrp.Sparklines.RemoveAt(i);
                                continue;
                            }
                            else if (spk.LocationRowIndex > iEndRowIndex)
                            {
                                iNewIndex = spk.LocationRowIndex - iNumberOfRows;
                                spk.LocationRowIndex = iNewIndex;
                            }
                            // no else because there's nothing done

                            // process only if the data source is on the currently selected worksheet
                            if (spk.WorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                if (StartRowIndex <= spk.StartRowIndex && spk.EndRowIndex <= iEndRowIndex)
                                {
                                    // the data source is completely within delete range
                                    // Excel 2010 keeps the WorksheetExtension, but I'm gonna just delete the whole thing.
                                    spkgrp.Sparklines.RemoveAt(i);
                                    continue;
                                }
                                else
                                {
                                    this.DeleteRowColumnIndexDelta(StartRowIndex, iEndRowIndex, iNumberOfRows, ref spk.StartRowIndex, ref spk.EndRowIndex);
                                }
                            }

                            spkgrp.Sparklines[i] = spk;
                        }
                    }
                }
                #endregion
            }

            return result;
        }

        /// <summary>
        /// Clear all cell content within specified rows. If the top-left cell of a merged cell is within specified rows, the merged cell content is also cleared.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <returns>True if content has been cleared. False otherwise. If there are no content within specified rows, false is also returned.</returns>
        public bool ClearRowContent(int StartRowIndex, int EndRowIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;

            bool result = false;
            List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
            List<int> listColumnKeys;
            foreach (int rowkey in listRowKeys)
            {
                if (iStartRowIndex <= rowkey && rowkey <= iEndRowIndex)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        this.ClearCellContentData(rowkey, colkey);
                        result = true;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Indicates if the column has an existing style.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <returns>True if the column has an existing style. False otherwise.</returns>
        public bool HasColumnStyle(string ColumnName)
        {
            bool result = false;
            result = HasColumnStyle(SLTool.ToColumnIndex(ColumnName));

            return result;
        }

        /// <summary>
        /// Indicates if the column has an existing style.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>True if the column has an existing style. False otherwise.</returns>
        public bool HasColumnStyle(int ColumnIndex)
        {
            bool result = false;
            if (slws.ColumnProperties.ContainsKey(ColumnIndex))
            {
                SLColumnProperties cp = slws.ColumnProperties[ColumnIndex];
                if (cp.StyleIndex > 0)
                {
                    result = true;
                }
            }

            return result;
        }

        /// <summary>
        /// Get the column width. If the column doesn't have a width explicitly set, the default column width for the current worksheet is returned.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <returns>The column width.</returns>
        public double GetColumnWidth(string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            return GetColumnWidth(iColumnIndex);
        }

        /// <summary>
        /// Get the column width. If the column doesn't have a width explicitly set, the default column width for the current worksheet is returned.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The column width.</returns>
        public double GetColumnWidth(int ColumnIndex)
        {
            double fWidth = slws.SheetFormatProperties.DefaultColumnWidth;
            if (slws.ColumnProperties.ContainsKey(ColumnIndex))
            {
                SLColumnProperties cp = slws.ColumnProperties[ColumnIndex];
                if (cp.HasWidth)
                {
                    fWidth = cp.Width;
                }
            }

            return fWidth;
        }

        /// <summary>
        /// Set the column width.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <param name="ColumnWidth">The column width.</param>
        /// <returns>True if the column name is valid. False otherwise.</returns>
        public bool SetColumnWidth(string ColumnName, double ColumnWidth)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            return SetColumnWidth(iColumnIndex, iColumnIndex, ColumnWidth);
        }

        /// <summary>
        /// Set the column width.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="ColumnWidth">The column width.</param>
        /// <returns>True if the column index is valid. False otherwise.</returns>
        public bool SetColumnWidth(int ColumnIndex, double ColumnWidth)
        {
            return SetColumnWidth(ColumnIndex, ColumnIndex, ColumnWidth);
        }

        /// <summary>
        /// Set the column width for a range of columns.
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column.</param>
        /// <param name="EndColumnName">The column name of the end column.</param>
        /// <param name="ColumnWidth">The column width.</param>
        /// <returns>True if the column names are valid. False otherwise.</returns>
        public bool SetColumnWidth(string StartColumnName, string EndColumnName, double ColumnWidth)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            return SetColumnWidth(iStartColumnIndex, iEndColumnIndex, ColumnWidth);
        }

        /// <summary>
        /// Set the column width for a range of columns.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        /// <param name="ColumnWidth">The column width.</param>
        /// <returns>True if the column indices are valid. False otherwise.</returns>
        public bool SetColumnWidth(int StartColumnIndex, int EndColumnIndex, double ColumnWidth)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            bool result = false;
            if (iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                result = true;
                int i = 0;
                SLColumnProperties cp;
                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    if (slws.ColumnProperties.ContainsKey(i))
                    {
                        cp = slws.ColumnProperties[i];
                        cp.Width = ColumnWidth;
                        slws.ColumnProperties[i] = cp;
                    }
                    else
                    {
                        cp = new SLColumnProperties(SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize);
                        cp.Width = ColumnWidth;
                        slws.ColumnProperties.Add(i, cp);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Automatically fit column width according to cell contents.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        public void AutoFitColumn(string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            this.AutoFitColumn(iColumnIndex, iColumnIndex, SLConstants.MaximumColumnWidth);
        }

        /// <summary>
        /// Automatically fit column width according to cell contents.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <param name="MaximumColumnWidth">The maximum column width in number of characters.</param>
        public void AutoFitColumn(string ColumnName, double MaximumColumnWidth)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            this.AutoFitColumn(iColumnIndex, iColumnIndex, MaximumColumnWidth);
        }

        /// <summary>
        /// Automatically fit column width according to cell contents.
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column.</param>
        /// <param name="EndColumnName">The column name of the end column.</param>
        public void AutoFitColumn(string StartColumnName, string EndColumnName)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            this.AutoFitColumn(iStartColumnIndex, iEndColumnIndex, SLConstants.MaximumColumnWidth);
        }

        /// <summary>
        /// Automatically fit column width according to cell contents.
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column.</param>
        /// <param name="EndColumnName">The column name of the end column.</param>
        /// <param name="MaximumColumnWidth">The maximum column width in number of characters.</param>
        public void AutoFitColumn(string StartColumnName, string EndColumnName, double MaximumColumnWidth)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            this.AutoFitColumn(iStartColumnIndex, iEndColumnIndex, MaximumColumnWidth);
        }

        /// <summary>
        /// Automatically fit column width according to cell contents.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        public void AutoFitColumn(int ColumnIndex)
        {
            this.AutoFitColumn(ColumnIndex, ColumnIndex, SLConstants.MaximumColumnWidth);
        }

        /// <summary>
        /// Automatically fit column width according to cell contents.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="MaximumColumnWidth">The maximum column width in number of characters.</param>
        public void AutoFitColumn(int ColumnIndex, double MaximumColumnWidth)
        {
            this.AutoFitColumn(ColumnIndex, ColumnIndex, MaximumColumnWidth);
        }

        /// <summary>
        /// Automatically fit column width according to cell contents.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        public void AutoFitColumn(int StartColumnIndex, int EndColumnIndex)
        {
            this.AutoFitColumn(StartColumnIndex, EndColumnIndex, SLConstants.MaximumColumnWidth);
        }

        /// <summary>
        /// Automatically fit column width according to cell contents.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        /// <param name="MaximumColumnWidth">The maximum column width in number of characters.</param>
        public void AutoFitColumn(int StartColumnIndex, int EndColumnIndex, double MaximumColumnWidth)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iStartColumnIndex > SLConstants.ColumnLimit) iStartColumnIndex = SLConstants.ColumnLimit;
            if (iEndColumnIndex < 1) iEndColumnIndex = 1;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            if (MaximumColumnWidth > SLConstants.MaximumColumnWidth) MaximumColumnWidth = SLConstants.MaximumColumnWidth;
            // this is taken from SLColumnProperties...
            int iWholeNumber = Convert.ToInt32(Math.Truncate(MaximumColumnWidth));
            double fStepRemainder = MaximumColumnWidth - (double)iWholeNumber;
            int iStep = 0;
            for (iStep = SimpleTheme.listColumnStepSize.Count - 1; iStep >= 0; --iStep)
            {
                if (fStepRemainder > SimpleTheme.listColumnStepSize[iStep]) break;
            }
            if (iStep < 0) iStep = 0;
            int iMaximumPixelLength = iWholeNumber * (SimpleTheme.ThemeMaxDigitWidth - 1) + iStep;

            Dictionary<int, int> pixellength = this.AutoFitRowColumn(false, iStartColumnIndex, iEndColumnIndex, iMaximumPixelLength);

            SLColumnProperties cp;
            double fColumnWidth;
            int iPixelLength;
            double fWholeNumber;
            double fRemainder;
            foreach (int pixlenpt in pixellength.Keys)
            {
                iPixelLength = pixellength[pixlenpt];
                if (iPixelLength > 0)
                {
                    fWholeNumber = (double)(iPixelLength / (SimpleTheme.ThemeMaxDigitWidth - 1));
                    fRemainder = (double)(iPixelLength % (SimpleTheme.ThemeMaxDigitWidth - 1));
                    fRemainder = fRemainder / (double)(SimpleTheme.ThemeMaxDigitWidth - 1);
                    // we'll leave it to the algorithm within SLColumnProperties.Width to handle
                    // the actual column width refitting.
                    fColumnWidth = fWholeNumber + fRemainder;
                    if (slws.ColumnProperties.ContainsKey(pixlenpt))
                    {
                        cp = slws.ColumnProperties[pixlenpt];
                        cp.Width = fColumnWidth;
                        cp.BestFit = true;
                        slws.ColumnProperties[pixlenpt] = cp.Clone();
                    }
                    else
                    {
                        cp = new SLColumnProperties(SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize);
                        cp.Width = fColumnWidth;
                        cp.BestFit = true;
                        slws.ColumnProperties[pixlenpt] = cp.Clone();
                    }
                }
                // else we don't have to do anything
            }
        }

        /// <summary>
        /// Indicates if the column is hidden.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <returns>True if the column is hidden. False otherwise.</returns>
        public bool IsColumnHidden(string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            return IsColumnHidden(iColumnIndex);
        }

        /// <summary>
        /// Indicates if the column is hidden.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>True if the column is hidden. False otherwise.</returns>
        public bool IsColumnHidden(int ColumnIndex)
        {
            bool result = false;
            if (slws.ColumnProperties.ContainsKey(ColumnIndex))
            {
                SLColumnProperties cp = slws.ColumnProperties[ColumnIndex];
                result = cp.Hidden;
            }

            return result;
        }

        private bool ToggleColumnHidden(int StartColumnIndex, int EndColumnIndex, bool Hidden)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            bool result = false;
            if (iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                result = true;
                int i = 0;
                SLColumnProperties cp;
                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    if (slws.ColumnProperties.ContainsKey(i))
                    {
                        cp = slws.ColumnProperties[i];
                        cp.Hidden = Hidden;
                        slws.ColumnProperties[i] = cp;
                    }
                    else
                    {
                        cp = new SLColumnProperties(SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize);
                        cp.Hidden = Hidden;
                        slws.ColumnProperties.Add(i, cp);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Hide the column.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <returns>True if the column name is valid. False otherwise.</returns>
        public bool HideColumn(string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            return HideColumn(iColumnIndex);
        }

        /// <summary>
        /// Hide the column.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>True if the column index is valid. False otherwise.</returns>
        public bool HideColumn(int ColumnIndex)
        {
            return ToggleColumnHidden(ColumnIndex, ColumnIndex, true);
        }

        /// <summary>
        /// Hide a range of columns.
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column.</param>
        /// <param name="EndColumnName">The column name of the end column.</param>
        /// <returns>True if the column names are valid. False otherwise.</returns>
        public bool HideColumn(string StartColumnName, string EndColumnName)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            return HideColumn(iStartColumnIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Hide a range of columns.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        /// <returns>True if the column indices are valid. False otherwise.</returns>
        public bool HideColumn(int StartColumnIndex, int EndColumnIndex)
        {
            return ToggleColumnHidden(StartColumnIndex, EndColumnIndex, true);
        }

        /// <summary>
        /// Unhide the column.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <returns>True if the column name is valid. False otherwise.</returns>
        public bool UnhideColumn(string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            return UnhideColumn(iColumnIndex);
        }

        /// <summary>
        /// Unhide the column.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>True if the column index is valid. False otherwise.</returns>
        public bool UnhideColumn(int ColumnIndex)
        {
            return ToggleColumnHidden(ColumnIndex, ColumnIndex, false);
        }

        /// <summary>
        /// Unhide a range of columns.
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column.</param>
        /// <param name="EndColumnName">The column name of the end column.</param>
        /// <returns>True if the column names are valid. False otherwise.</returns>
        public bool UnhideColumn(string StartColumnName, string EndColumnName)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            return UnhideColumn(iStartColumnIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Unhide a range of columns.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        /// <returns>True if the column indices are valid. False otherwise.</returns>
        public bool UnhideColumn(int StartColumnIndex, int EndColumnIndex)
        {
            return ToggleColumnHidden(StartColumnIndex, EndColumnIndex, false);
        }

        /// <summary>
        /// Indicates if the column is showing phonetic information.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <returns>True if the column is showing phonetic information. False otherwise.</returns>
        public bool IsColumnShowingPhonetic(string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            return IsColumnShowingPhonetic(iColumnIndex);
        }

        /// <summary>
        /// Indicates if the column is showing phonetic information.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>True if the column is showing phonetic information. False otherwise.</returns>
        public bool IsColumnShowingPhonetic(int ColumnIndex)
        {
            bool result = false;
            if (slws.ColumnProperties.ContainsKey(ColumnIndex))
            {
                SLColumnProperties cp = slws.ColumnProperties[ColumnIndex];
                result = cp.Phonetic;
            }

            return result;
        }

        /// <summary>
        /// Set the show phonetic property for the column.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <param name="ShowPhonetic">True if the column should show phonetic information. False otherwise.</param>
        /// <returns>True if the column name is valid. False otherwise.</returns>
        public bool SetColumnShowPhonetic(string ColumnName, bool ShowPhonetic)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            return SetColumnShowPhonetic(iColumnIndex, iColumnIndex, ShowPhonetic);
        }

        /// <summary>
        /// Set the show phonetic property for the column.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="ShowPhonetic">True if the column should show phonetic information. False otherwise.</param>
        /// <returns>True if the column index is valid. False otherwise.</returns>
        public bool SetColumnShowPhonetic(int ColumnIndex, bool ShowPhonetic)
        {
            return SetColumnShowPhonetic(ColumnIndex, ColumnIndex, ShowPhonetic);
        }

        /// <summary>
        /// Set the show phonetic property for a range of columns.
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column.</param>
        /// <param name="EndColumnName">The column name of the end column.</param>
        /// <param name="ShowPhonetic">True if the columns should show phonetic information. False otherwise.</param>
        /// <returns>True if the column names are valid. False otherwise.</returns>
        public bool SetColumnShowPhonetic(string StartColumnName, string EndColumnName, bool ShowPhonetic)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            return SetColumnShowPhonetic(iStartColumnIndex, iEndColumnIndex, ShowPhonetic);
        }

        /// <summary>
        /// Set the show phonetic property for a range of columns.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        /// <param name="ShowPhonetic">True if the columns should show phonetic information. False otherwise.</param>
        /// <returns>True if the column indices are valid. False otherwise.</returns>
        public bool SetColumnShowPhonetic(int StartColumnIndex, int EndColumnIndex, bool ShowPhonetic)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            bool result = false;
            if (iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                result = true;
                int i = 0;
                SLColumnProperties cp;
                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    if (slws.ColumnProperties.ContainsKey(i))
                    {
                        cp = slws.ColumnProperties[i];
                        cp.Phonetic = ShowPhonetic;
                        slws.ColumnProperties[i] = cp;
                    }
                    else
                    {
                        cp = new SLColumnProperties(SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize);
                        cp.Phonetic = ShowPhonetic;
                        slws.ColumnProperties.Add(i, cp);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Copy one column to another column.
        /// </summary>
        /// <param name="ColumnName">The column name of the column to be copied from.</param>
        /// <param name="AnchorColumnName">The column name of the column to be copied to.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumn(string ColumnName, string AnchorColumnName)
        {
            int iColumnIndex = -1;
            int iAnchorColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);
            iAnchorColumnIndex = SLTool.ToColumnIndex(AnchorColumnName);

            return CopyColumn(iColumnIndex, iColumnIndex, iAnchorColumnIndex, false);
        }

        /// <summary>
        /// Copy one column to another column.
        /// </summary>
        /// <param name="ColumnIndex">The column index of the column to be copied from.</param>
        /// <param name="AnchorColumnIndex">The column index of the column to be copied to.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumn(int ColumnIndex, int AnchorColumnIndex)
        {
            return CopyColumn(ColumnIndex, ColumnIndex, AnchorColumnIndex, false);
        }

        /// <summary>
        /// Copy one column to another column.
        /// </summary>
        /// <param name="ColumnName">The column name of the column to be copied from.</param>
        /// <param name="AnchorColumnName">The column name of the column to be copied to.</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumn(string ColumnName, string AnchorColumnName, bool ToCut)
        {
            int iColumnIndex = -1;
            int iAnchorColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);
            iAnchorColumnIndex = SLTool.ToColumnIndex(AnchorColumnName);

            return CopyColumn(iColumnIndex, iColumnIndex, iAnchorColumnIndex, ToCut);
        }

        /// <summary>
        /// Copy one column to another column.
        /// </summary>
        /// <param name="ColumnIndex">The column index of the column to be copied from.</param>
        /// <param name="AnchorColumnIndex">The column index of the column to be copied to.</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumn(int ColumnIndex, int AnchorColumnIndex, bool ToCut)
        {
            return CopyColumn(ColumnIndex, ColumnIndex, AnchorColumnIndex, ToCut);
        }

        /// <summary>
        /// Copy a range of columns to another range, given the anchor column of the destination range (left-most column).
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column of the column range. This is typically the left-most column.</param>
        /// <param name="EndColumnName">The column name of the end column of the column range. This is typically the right-most column.</param>
        /// <param name="AnchorColumnName">The column name of the anchor column.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumn(string StartColumnName, string EndColumnName, string AnchorColumnName)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            int iAnchorColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);
            iAnchorColumnIndex = SLTool.ToColumnIndex(AnchorColumnName);

            return CopyColumn(iStartColumnIndex, iEndColumnIndex, iAnchorColumnIndex, false);
        }

        /// <summary>
        /// Copy a range of columns to another range, given the anchor column of the destination range (left-most column).
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column of the column range. This is typically the left-most column.</param>
        /// <param name="EndColumnIndex">The column index of the end column of the column range. This is typically the right-most column.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor column.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumn(int StartColumnIndex, int EndColumnIndex, int AnchorColumnIndex)
        {
            return CopyColumn(StartColumnIndex, EndColumnIndex, AnchorColumnIndex, false);
        }

        /// <summary>
        /// Copy a range of columns to another range, given the anchor column of the destination range (left-most column).
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column of the column range. This is typically the left-most column.</param>
        /// <param name="EndColumnName">The column name of the end column of the column range. This is typically the right-most column.</param>
        /// <param name="AnchorColumnName">The column name of the anchor column.</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumn(string StartColumnName, string EndColumnName, string AnchorColumnName, bool ToCut)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            int iAnchorColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);
            iAnchorColumnIndex = SLTool.ToColumnIndex(AnchorColumnName);

            return CopyColumn(iStartColumnIndex, iEndColumnIndex, iAnchorColumnIndex, ToCut);
        }

        /// <summary>
        /// Copy a range of columns to another range, given the anchor column of the destination range (left-most column).
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column of the column range. This is typically the left-most column.</param>
        /// <param name="EndColumnIndex">The column index of the end column of the column range. This is typically the right-most column.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor column.</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumn(int StartColumnIndex, int EndColumnIndex, int AnchorColumnIndex, bool ToCut)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            bool result = false;
            if (iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit
                && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit
                && AnchorColumnIndex >= 1 && AnchorColumnIndex <= SLConstants.ColumnLimit
                && iStartColumnIndex != AnchorColumnIndex)
            {
                this.FlattenAllSharedCellFormula();

                result = true;

                int diff = AnchorColumnIndex - iStartColumnIndex;
                int i = 0;
                Dictionary<int, SLColumnProperties> cols = new Dictionary<int, SLColumnProperties>();
                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    if (slws.ColumnProperties.ContainsKey(i))
                    {
                        cols[i + diff] = slws.ColumnProperties[i].Clone();
                        if (ToCut)
                        {
                            slws.ColumnProperties.Remove(i);
                        }
                    }
                }

                int AnchorEndColumnIndex = AnchorColumnIndex + iEndColumnIndex - iStartColumnIndex;
                // removing columns within destination "paste" operation
                List<int> colkeys = slws.ColumnProperties.Keys.ToList<int>();
                foreach (int ckey in colkeys)
                {
                    if (ckey >= AnchorColumnIndex && ckey <= AnchorEndColumnIndex)
                    {
                        slws.ColumnProperties.Remove(ckey);
                    }
                }

                foreach (var key in cols.Keys)
                {
                    slws.ColumnProperties[key] = cols[key].Clone();
                }

                SLCellWarehouse cells = new SLCellWarehouse();
                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                List<int> listColumnKeys;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        if (colkey >= iStartColumnIndex && colkey <= iEndColumnIndex)
                        {
                            cells.SetValue(rowkey, colkey + diff, slws.CellWarehouse.Cells[rowkey][colkey]);
                            if (ToCut)
                            {
                                slws.CellWarehouse.Remove(rowkey, colkey);
                            }
                        }
                    }
                }

                listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        // any cell within destination "paste" operation is taken out
                        if (colkey >= AnchorColumnIndex && colkey <= AnchorEndColumnIndex)
                        {
                            slws.CellWarehouse.Remove(rowkey, colkey);
                        }
                    }
                }

                int iNumberOfColumns = iEndColumnIndex - iStartColumnIndex + 1;
                if (AnchorColumnIndex <= iStartColumnIndex) iNumberOfColumns = -iNumberOfColumns;

                SLCell c;
                listRowKeys = cells.Cells.Keys.ToList<int>();
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = cells.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        c = cells.Cells[rowkey][colkey].Clone();
                        //this.ProcessCellFormulaDelta(ref c, -1, 0, AnchorColumnIndex, iNumberOfColumns);
                        this.ProcessCellFormulaDelta(ref c, false, 0, StartColumnIndex, 0, StartColumnIndex + iNumberOfColumns, false, false, false, 0, 0);
                        slws.CellWarehouse.SetValue(rowkey, colkey, c);
                    }
                }

                // TODO: tables!

                // cutting and pasting into a region with merged cells unmerges the existing merged cells
                // copying and pasting into a region with merged cells leaves existing merged cells alone.
                // Why does Excel do that? Don't know.
                // Will just standardise to leaving existing merged cells alone.
                List<SLMergeCell> mca = this.GetWorksheetMergeCells();
                foreach (SLMergeCell mc in mca)
                {
                    if (mc.StartColumnIndex >= iStartColumnIndex && mc.EndColumnIndex <= iEndColumnIndex)
                    {
                        if (ToCut)
                        {
                            slws.MergeCells.Remove(mc);
                        }
                        this.MergeWorksheetCells(mc.StartRowIndex, mc.StartColumnIndex + diff, mc.EndRowIndex, mc.EndColumnIndex + diff);
                    }
                }

                #region Calculation cells
                if (slwb.CalculationCells.Count > 0)
                {
                    // I don't know enough to fiddle with calculation chains. So I'm going to ignore it.
                    slwb.CalculationCells.Clear();

                    //List<int> listToDelete = new List<int>();
                    //int iColumnLimit = AnchorColumnIndex + iStartColumnIndex - iEndColumnIndex;
                    //for (i = 0; i < slwb.CalculationCells.Count; ++i)
                    //{
                    //    if (slwb.CalculationCells[i].SheetId == giSelectedWorksheetID)
                    //    {
                    //        if (ToCut && slwb.CalculationCells[i].ColumnIndex >= iStartColumnIndex && slwb.CalculationCells[i].ColumnIndex <= iEndColumnIndex)
                    //        {
                    //            // just remove because recalculation of cell references is too complicated...
                    //            if (!listToDelete.Contains(i)) listToDelete.Add(i);
                    //        }

                    //        if (slwb.CalculationCells[i].ColumnIndex >= AnchorColumnIndex && slwb.CalculationCells[i].ColumnIndex <= iColumnLimit)
                    //        {
                    //            // existing calculation cell lies within destination "paste" operation
                    //            if (!listToDelete.Contains(i)) listToDelete.Add(i);
                    //        }
                    //    }
                    //}

                    //for (i = listToDelete.Count - 1; i >= 0; --i)
                    //{
                    //    slwb.CalculationCells.RemoveAt(listToDelete[i]);
                    //}
                }
                #endregion

                // defined names is hard to calculate...
                // need to check the row and column indices based on the cell references within.
            }

            return result;
        }

        /// <summary>
        /// Get the group level of the column.
        /// </summary>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <returns>The group level (between 0 and 7, both inclusive). The default is 0.</returns>
        public int GetColumnGroupLevel(string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            return GetColumnGroupLevel(iColumnIndex);
        }

        /// <summary>
        /// Get the group level of the column.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The group level (between 0 and 7, both inclusive). The default is 0.</returns>
        public int GetColumnGroupLevel(int ColumnIndex)
        {
            int iGroupLevel = 0;
            if (slws.ColumnProperties.ContainsKey(ColumnIndex))
            {
                iGroupLevel = slws.ColumnProperties[ColumnIndex].OutlineLevel;
            }

            return iGroupLevel;
        }

        /// <summary>
        /// Group columns.
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column of the column range. This is typically the left-most column.</param>
        /// <param name="EndColumnName">The column name of the end column of the column range. This is typically the right-most column.</param>
        public void GroupColumns(string StartColumnName, string EndColumnName)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            GroupColumns(iStartColumnIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Group columns.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        public void GroupColumns(int StartColumnIndex, int EndColumnIndex)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            // I haven't personally checked this, but there's a collapsing -/+ box on the column
            // just right of the grouped columns. This implies the very very last column that can be
            // grouped is the (column limit - 1)th column, because (column limit)th column will have that
            // collapsing box.
            if (iEndColumnIndex >= SLConstants.RowLimit) iEndColumnIndex = SLConstants.ColumnLimit - 1;
            // there's nothing to group...
            if (iStartColumnIndex > iEndColumnIndex) return;

            if (iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                int i = 0;
                SLColumnProperties cp;
                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    if (slws.ColumnProperties.ContainsKey(i))
                    {
                        cp = slws.ColumnProperties[i];
                        // Excel supports only levels 0 to 7
                        if (cp.OutlineLevel < 7) ++cp.OutlineLevel;
                        slws.ColumnProperties[i] = cp.Clone();
                    }
                    else
                    {
                        cp = new SLColumnProperties(SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize);
                        cp.OutlineLevel = 1;
                        slws.ColumnProperties[i] = cp.Clone();
                    }
                }
            }
        }

        /// <summary>
        /// Ungroup columns.
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column of the column range. This is typically the left-most column.</param>
        /// <param name="EndColumnName">The column name of the end column of the column range. This is typically the right-most column.</param>
        public void UngroupColumns(string StartColumnName, string EndColumnName)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            UngroupColumns(iStartColumnIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Ungroup columns.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        public void UngroupColumns(int StartColumnIndex, int EndColumnIndex)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            // the following algorithm is not guaranteed to work in all cases.
            // The data is sort of loosely linked together with no guarantee that they
            // all make sense together. If you use Excel, then the internal data is sort of
            // guaranteed to make sense together, but anyone can make an Open XML spreadsheet
            // with possibly invalid-looking data. Maybe Excel will accept it, maybe not.

            if (iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                SLColumnProperties cp;
                byte byCurrentOutlineLevel;
                int i;

                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    if (slws.ColumnProperties.ContainsKey(i))
                    {
                        cp = slws.ColumnProperties[i];
                        if (cp.OutlineLevel > 0) --cp.OutlineLevel;
                        slws.ColumnProperties[i] = cp.Clone();

                        // if after ungrouping, the outline level is the same as the next
                        // one and the next one is collapsed, then we probably reached the
                        // end of the group and we uncollapse the thing. It's not so much
                        // an uncollapse but an indication to tell the application
                        // (read: Excel) not to choke on missing groups with a collapse command.
                        byCurrentOutlineLevel = cp.OutlineLevel;
                        if (slws.ColumnProperties.ContainsKey(i + 1))
                        {
                            cp = slws.ColumnProperties[i + 1];
                            if (cp.OutlineLevel == byCurrentOutlineLevel && cp.Collapsed)
                            {
                                cp.Collapsed = false;
                                slws.ColumnProperties[i + 1] = cp.Clone();
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Collapse a group of columns.
        /// </summary>
        /// <param name="ColumnName">The column name (such as "A1") of the column just after the group of columns you want to collapse. For example, this will be column E if columns B to D are grouped.</param>
        public void CollapseColumns(string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            CollapseColumns(iColumnIndex);
        }

        /// <summary>
        /// Collapse a group of columns.
        /// </summary>
        /// <param name="ColumnIndex">The column index of the column just after the group of columns you want to collapse. For example, this will be column 5 if columns 2 to 4 are grouped.</param>
        public void CollapseColumns(int ColumnIndex)
        {
            if (ColumnIndex < 1 || ColumnIndex > SLConstants.ColumnLimit) return;

            // the following algorithm is not guaranteed to work in all cases.
            // The data is sort of loosely linked together with no guarantee that they
            // all make sense together. If you use Excel, then the internal data is sort of
            // guaranteed to make sense together, but anyone can make an Open XML spreadsheet
            // with possibly invalid-looking data. Maybe Excel will accept it, maybe not.

            SLColumnProperties cp;
            byte byCurrentOutlineLevel = 0;
            if (slws.ColumnProperties.ContainsKey(ColumnIndex))
            {
                cp = slws.ColumnProperties[ColumnIndex];
                byCurrentOutlineLevel = cp.OutlineLevel;
            }

            bool bFound = false;
            int i;

            for (i = ColumnIndex - 1; i >= 1; --i)
            {
                if (slws.ColumnProperties.ContainsKey(i))
                {
                    cp = slws.ColumnProperties[i];
                    if (cp.OutlineLevel > byCurrentOutlineLevel)
                    {
                        bFound = true;
                        cp.Hidden = true;
                        slws.ColumnProperties[i] = cp.Clone();
                    }
                    else break;
                }
                else break;
            }

            if (bFound)
            {
                if (slws.ColumnProperties.ContainsKey(ColumnIndex))
                {
                    cp = slws.ColumnProperties[ColumnIndex];
                    cp.Collapsed = true;
                    slws.ColumnProperties[ColumnIndex] = cp.Clone();
                }
                else
                {
                    cp = new SLColumnProperties(SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize);
                    cp.Collapsed = true;
                    slws.ColumnProperties[ColumnIndex] = cp.Clone();
                }
            }
        }

        /// <summary>
        /// Expand a group of columns.
        /// </summary>
        /// <param name="ColumnName">The column name (such as "A1") of the column just after the group of columns you want to expand. For example, this will be column E if columns B to D are grouped.</param>
        public void ExpandColumns(string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            ExpandColumns(iColumnIndex);
        }

        /// <summary>
        /// Expand a group of columns.
        /// </summary>
        /// <param name="ColumnIndex">The column index of the column just after the group of columns you want to expand. For example, this will be column 5 if columns 2 to 4 are grouped.</param>
        public void ExpandColumns(int ColumnIndex)
        {
            if (ColumnIndex < 1 || ColumnIndex > SLConstants.ColumnLimit) return;

            // the following algorithm is not guaranteed to work in all cases.
            // The data is sort of loosely linked together with no guarantee that they
            // all make sense together. If you use Excel, then the internal data is sort of
            // guaranteed to make sense together, but anyone can make an Open XML spreadsheet
            // with possibly invalid-looking data. Maybe Excel will accept it, maybe not.

            if (slws.ColumnProperties.ContainsKey(ColumnIndex))
            {
                SLColumnProperties cp = slws.ColumnProperties[ColumnIndex];
                // no point if it's not the collapsing -/+ box
                if (cp.Collapsed)
                {
                    if (cp.Hidden)
                    {
                        // if it's hidden, it's probably because it and it's associated
                        // group is hidden behind another group. So we don't show the rest
                        // of the group.
                        cp.Collapsed = false;
                        slws.ColumnProperties[ColumnIndex] = cp.Clone();
                        // Of course I don't really know that for sure. Hence the "probably".
                    }
                    else
                    {
                        cp.Collapsed = false;
                        slws.ColumnProperties[ColumnIndex] = cp.Clone();

                        byte byCurrentOutlineLevel = cp.OutlineLevel;
                        int i;
                        for (i = ColumnIndex - 1; i >= 1; --i)
                        {
                            if (slws.ColumnProperties.ContainsKey(i))
                            {
                                cp = slws.ColumnProperties[i];
                                if (cp.OutlineLevel > byCurrentOutlineLevel)
                                {
                                    cp.Hidden = false;
                                    slws.ColumnProperties[i] = cp.Clone();
                                }
                                else break;
                            }
                            else break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Insert one or more columns.
        /// </summary>
        /// <param name="StartColumnName">Additional columns will be inserted at this column.</param>
        /// <param name="NumberOfColumns">Number of columns to insert.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertColumn(string StartColumnName, int NumberOfColumns)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(StartColumnName);

            return InsertColumn(iColumnIndex, NumberOfColumns);
        }

        /// <summary>
        /// Insert one or more columns.
        /// </summary>
        /// <param name="StartColumnIndex">Additional columns will be inserted at this column index.</param>
        /// <param name="NumberOfColumns">Number of columns to insert.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertColumn(int StartColumnIndex, int NumberOfColumns)
        {
            if (NumberOfColumns < 1) return false;

            bool result = false;
            if (StartColumnIndex >= 1 && StartColumnIndex <= SLConstants.ColumnLimit)
            {
                this.FlattenAllSharedCellFormula();

                result = true;
                int i = 0, iNewIndex = 0;

                int index = 0;
                //int iRowIndex = -1, iColumnIndex = -1;
                //int iRowIndex2 = -1, iColumnIndex2 = -1;
                int iColumnIndex = -1;
                int iColumnIndex2 = -1;

                #region Column properties
                SLColumnProperties cp;
                List<int> listColumnIndex = slws.ColumnProperties.Keys.ToList<int>();
                // this sorting in descending order is crucial!
                // we move the data from after the insert range to their new reference keys
                // first, then we put in the new data, which will then have no data
                // key collision.
                listColumnIndex.Sort();
                listColumnIndex.Reverse();

                for (i = 0; i < listColumnIndex.Count; ++i)
                {
                    index = listColumnIndex[i];
                    if (index >= StartColumnIndex)
                    {
                        cp = slws.ColumnProperties[index];
                        slws.ColumnProperties.Remove(index);
                        iNewIndex = index + NumberOfColumns;
                        // if the new column is right of right-side limit of the worksheet,
                        // then it disappears into the ether...
                        if (iNewIndex <= SLConstants.ColumnLimit)
                        {
                            slws.ColumnProperties[iNewIndex] = cp.Clone();
                        }
                    }
                    else
                    {
                        // the columns before the start column are unaffected by the insertion.
                        // Because it's sorted in descending order, we can just break out.
                        break;
                    }
                }
                #endregion

                #region Cell data
                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                // 2 April 2017: Need to sort or the row keys are "random" (depending on
                // how the OS assign memory and order and such)
                // This made the rows "random"
                listRowKeys.Sort();
                List<int> listColumnKeys;

                SLCell c;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    listColumnKeys.Sort();
                    // 5 July 2019: Need to reverse order so that columns are inserted correctly
                    // (thanks Azamat S. !)
                    listColumnKeys.Reverse();
                    foreach (int colkey in listColumnKeys)
                    {
                        c = slws.CellWarehouse.Cells[rowkey][colkey].Clone();
                        //this.ProcessCellFormulaDelta(ref c, -1, 0, StartColumnIndex, NumberOfColumns);
                        this.ProcessCellFormulaDelta(ref c, false, 0, StartColumnIndex, 0, StartColumnIndex + NumberOfColumns, true, true, false, 0, 0);

                        if (colkey >= StartColumnIndex)
                        {
                            slws.CellWarehouse.Remove(rowkey, colkey);
                            iNewIndex = colkey + NumberOfColumns;
                            if (iNewIndex <= SLConstants.ColumnLimit)
                            {
                                slws.CellWarehouse.SetValue(rowkey, iNewIndex, c);
                            }
                        }
                        else
                        {
                            slws.CellWarehouse.SetValue(rowkey, colkey, c);
                        }
                    }
                }

                #region Cell comments
                List<SLCellPoint> listCellRefKeys = slws.Comments.Keys.ToList<SLCellPoint>();
                // this sorting in descending order is crucial!
                listCellRefKeys.Sort(new SLCellReferencePointComparer());
                listCellRefKeys.Reverse();

                SLCellPoint pt;
                SLComment comm;
                for (i = 0; i < listCellRefKeys.Count; ++i)
                {
                    pt = listCellRefKeys[i];
                    comm = slws.Comments[pt];
                    if (pt.ColumnIndex >= StartColumnIndex)
                    {
                        slws.Comments.Remove(pt);
                        iNewIndex = pt.ColumnIndex + NumberOfColumns;
                        if (iNewIndex <= SLConstants.ColumnLimit)
                        {
                            slws.Comments[new SLCellPoint(pt.RowIndex, iNewIndex)] = comm;
                        }
                    }
                    // no else because there's nothing done
                }
                #endregion

                #endregion

                // the tables part has to be after the cell data part because we need
                // the cells to be correctly adjusted first. The insertion of new columns
                // also means updating some cells for the column header names.

                // Excel doesn't seem to allow inserting/deleting columns in certain cases
                // when 2 (or more) tables overlap each other vertically (meaning one above the other).
                // In particular, when the insert/delete range overlaps an existing column
                // in 2 (or more) tables.
                // The algorithm below works fine, meaning the resulting spreadsheet doesn't
                // have errors, so not sure why Excel doesn't allow it.
                #region Tables
                if (slws.Tables.Count > 0)
                {
                    int iNewID = 0;
                    string sNewColumnName = string.Empty;
                    int iCount = 0;

                    foreach (SLTable t in slws.Tables)
                    {
                        iColumnIndex = t.StartColumnIndex;
                        iColumnIndex2 = t.EndColumnIndex;
                        // need to modify table columns if the start column index is between the
                        // table reference columns, inclusive of the end column. If the start column index
                        // is the same as (or before) the start column of the table, then the whole
                        // table is shifted, so no modification is needed.
                        if (iColumnIndex < StartColumnIndex && StartColumnIndex <= iColumnIndex2)
                        {
                            for (i = 0; i < NumberOfColumns; ++i)
                            {
                                // the new ID and column name should be found long before the
                                // column limit is hit. Unless the table is unusually large...
                                for (iNewID = 1; iNewID <= SLConstants.ColumnLimit; ++iNewID)
                                {
                                    sNewColumnName = string.Format("Column{0}", iNewID);
                                    iCount = t.TableColumns.Count(n => n.Name.Equals(sNewColumnName, StringComparison.OrdinalIgnoreCase));
                                    if (iCount == 0) break;
                                }

                                for (iNewID = 1; iNewID <= SLConstants.ColumnLimit; ++iNewID)
                                {
                                    iCount = t.TableColumns.Count(n => n.Id == iNewID);
                                    if (iCount == 0) break;
                                }

                                if (t.HeaderRowCount > 0)
                                {
                                    iNewIndex = StartColumnIndex + i;
                                    if (iNewIndex > SLConstants.ColumnLimit) iNewIndex = SLConstants.ColumnLimit;
                                    this.SetCellValue(t.StartRowIndex, iNewIndex, sNewColumnName);
                                }

                                t.TableColumns.Insert(StartColumnIndex - iColumnIndex + i, new SLTableColumn() { Id = (uint)iNewID, Name = sNewColumnName });
                            }

                            // remove any extra columns that hang outside the worksheet after insertion
                            iCount = StartColumnIndex + NumberOfColumns - SLConstants.ColumnLimit;
                            for (i = 0; i < iCount; ++i)
                            {
                                // keep removing the last one
                                t.TableColumns.RemoveAt(t.TableColumns.Count - 1);
                            }
                        }

                        this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref iColumnIndex, ref iColumnIndex2);
                        if (iColumnIndex != t.StartColumnIndex || iColumnIndex2 != t.EndColumnIndex) t.IsNewTable = true;
                        t.StartColumnIndex = iColumnIndex;
                        t.EndColumnIndex = iColumnIndex2;

                        if (t.HasAutoFilter)
                        {
                            iColumnIndex = t.AutoFilter.StartColumnIndex;
                            iColumnIndex2 = t.AutoFilter.EndColumnIndex;
                            this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref iColumnIndex, ref iColumnIndex2);
                            if (iColumnIndex != t.AutoFilter.StartColumnIndex || iColumnIndex2 != t.AutoFilter.EndColumnIndex) t.IsNewTable = true;
                            t.AutoFilter.StartColumnIndex = iColumnIndex;
                            t.AutoFilter.EndColumnIndex = iColumnIndex2;

                            if (t.AutoFilter.HasSortState)
                            {
                                iColumnIndex = t.AutoFilter.SortState.StartColumnIndex;
                                iColumnIndex2 = t.AutoFilter.SortState.EndColumnIndex;
                                this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref iColumnIndex, ref iColumnIndex2);
                                if (iColumnIndex != t.AutoFilter.SortState.StartColumnIndex || iColumnIndex2 != t.AutoFilter.SortState.EndColumnIndex) t.IsNewTable = true;
                                t.AutoFilter.SortState.StartColumnIndex = iColumnIndex;
                                t.AutoFilter.SortState.EndColumnIndex = iColumnIndex2;
                            }
                        }

                        if (t.HasSortState)
                        {
                            iColumnIndex = t.SortState.StartColumnIndex;
                            iColumnIndex2 = t.SortState.EndColumnIndex;
                            this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref iColumnIndex, ref iColumnIndex2);
                            if (iColumnIndex != t.SortState.StartColumnIndex || iColumnIndex2 != t.SortState.EndColumnIndex) t.IsNewTable = true;
                            t.SortState.StartColumnIndex = iColumnIndex;
                            t.SortState.EndColumnIndex = iColumnIndex2;
                        }
                    }
                }
                #endregion

                #region Merge cells
                if (slws.MergeCells.Count > 0)
                {
                    SLMergeCell mc;
                    for (i = 0; i < slws.MergeCells.Count; ++i)
                    {
                        mc = slws.MergeCells[i];
                        this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref mc.iStartColumnIndex, ref mc.iEndColumnIndex);
                        slws.MergeCells[i] = mc;
                    }
                }
                #endregion

                #region Hyperlinks
                if (slws.Hyperlinks.Count > 0)
                {
                    SLHyperlink hl;
                    for (i = 0; i < slws.Hyperlinks.Count; ++i)
                    {
                        hl = slws.Hyperlinks[i];
                        iColumnIndex = hl.Reference.StartColumnIndex;
                        iColumnIndex2 = hl.Reference.EndColumnIndex;
                        this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref iColumnIndex, ref iColumnIndex2);
                        hl.Reference = new SLCellPointRange(hl.Reference.StartRowIndex, iColumnIndex, hl.Reference.EndRowIndex, iColumnIndex2);
                        slws.Hyperlinks[i] = hl.Clone();
                    }
                }
                #endregion

                #region Drawings
                if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
                {
                    WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                    if (wsp.DrawingsPart != null)
                    {
                        bool bFound = false;
                        Xdr.TwoCellAnchor tcaNew;
                        Xdr.OneCellAnchor ocaNew;
                        Xdr.EditAsValues vEditAs = Xdr.EditAsValues.Absolute;
                        List<OpenXmlElement> listoxe = new List<OpenXmlElement>();
                        int iIndex = 0;

                        DrawingsPart dp = wsp.DrawingsPart;
                        foreach (OpenXmlElement oxe in dp.WorksheetDrawing.ChildElements)
                        {
                            if (oxe is Xdr.TwoCellAnchor)
                            {
                                tcaNew = (Xdr.TwoCellAnchor)oxe.CloneNode(true);

                                if (tcaNew.EditAs == null) vEditAs = Xdr.EditAsValues.TwoCell;
                                else vEditAs = tcaNew.EditAs.Value;

                                if (vEditAs == Xdr.EditAsValues.TwoCell)
                                {
                                    if (tcaNew.FromMarker != null && tcaNew.FromMarker.ColumnId != null)
                                    {
                                        iIndex = Convert.ToInt32(tcaNew.FromMarker.ColumnId.Text);
                                        // the index is 0-based while the spreadsheet index is 1-based.
                                        if ((iIndex + 1) >= StartColumnIndex)
                                        {
                                            iIndex += NumberOfColumns;
                                            bFound = true;
                                            tcaNew.FromMarker.ColumnId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                        }
                                    }

                                    if (tcaNew.ToMarker != null && tcaNew.ToMarker.ColumnId != null)
                                    {
                                        iIndex = Convert.ToInt32(tcaNew.ToMarker.ColumnId.Text);
                                        // the index is 0-based while the spreadsheet index is 1-based.
                                        if ((iIndex + 1) >= StartColumnIndex)
                                        {
                                            iIndex += NumberOfColumns;
                                            bFound = true;
                                            tcaNew.ToMarker.ColumnId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                        }
                                    }
                                }
                                else if (vEditAs == Xdr.EditAsValues.OneCell)
                                {
                                    if (tcaNew.FromMarker != null && tcaNew.FromMarker.ColumnId != null)
                                    {
                                        iIndex = Convert.ToInt32(tcaNew.FromMarker.ColumnId.Text);
                                        // the index is 0-based while the spreadsheet index is 1-based.
                                        if ((iIndex + 1) >= StartColumnIndex)
                                        {
                                            iIndex += NumberOfColumns;
                                            bFound = true;
                                            tcaNew.FromMarker.ColumnId.Text = iIndex.ToString(CultureInfo.InvariantCulture);

                                            // if the from marker is moved, then the to marker has to move too
                                            if (tcaNew.ToMarker != null && tcaNew.ToMarker.ColumnId != null)
                                            {
                                                iIndex = Convert.ToInt32(tcaNew.ToMarker.ColumnId.Text);
                                                iIndex += NumberOfColumns;
                                                tcaNew.ToMarker.ColumnId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                            }
                                        }
                                    }
                                }

                                // Need to do for Transform for the child elements?

                                listoxe.Add(tcaNew.CloneNode(true));
                            }
                            else if (oxe is Xdr.OneCellAnchor)
                            {
                                ocaNew = (Xdr.OneCellAnchor)oxe.CloneNode(true);
                                if (ocaNew.FromMarker != null && ocaNew.FromMarker.ColumnId != null)
                                {
                                    iIndex = Convert.ToInt32(ocaNew.FromMarker.ColumnId.Text);
                                    // the index is 0-based while the spreadsheet index is 1-based.
                                    if ((iIndex + 1) >= StartColumnIndex)
                                    {
                                        iIndex += NumberOfColumns;
                                        bFound = true;
                                        ocaNew.FromMarker.ColumnId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                    }
                                }

                                // Need to do for Transform for the child elements?

                                listoxe.Add(ocaNew.CloneNode(true));
                            }
                            else
                            {
                                listoxe.Add(oxe.CloneNode(true));
                            }
                        }

                        if (bFound)
                        {
                            wsp.DrawingsPart.WorksheetDrawing.RemoveAllChildren();
                            foreach (OpenXmlElement oxe in listoxe)
                            {
                                wsp.DrawingsPart.WorksheetDrawing.Append(oxe.CloneNode(true));
                            }
                            wsp.DrawingsPart.WorksheetDrawing.Save();
                        }
                    }
                }
                #endregion

                // TODO: chart series references

                #region Calculation chain
                if (slwb.CalculationCells.Count > 0)
                {
                    // I don't know enough to fiddle with calculation chains. So I'm going to ignore it.
                    slwb.CalculationCells.Clear();

                    //foreach (SLCalculationCell cc in slwb.CalculationCells)
                    //{
                    //    if (cc.SheetId == giSelectedWorksheetID)
                    //    {
                    //        iColumnIndex = cc.ColumnIndex;
                    //        // don't need this but assign something anyway...
                    //        iColumnIndex2 = SLConstants.ColumnLimit;

                    //        this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref iColumnIndex, ref iColumnIndex2);
                    //        cc.ColumnIndex = iColumnIndex;
                    //    }
                    //}
                }
                #endregion

                #region Shared formulas
                if (slws.SharedCellFormulas.Count > 0)
                {
                    SLSharedCellFormula scf;
                    bool bToDelete;
                    List<uint> scfkeys = slws.SharedCellFormulas.Keys.ToList<uint>();
                    for (i = scfkeys.Count - 1; i >= 0; --i)
                    {
                        scf = slws.SharedCellFormulas[scfkeys[i]];
                        bToDelete = false;
                        // start from the end because we might be deleting
                        for (index = scf.Reference.Count - 1; index >= 0; --index)
                        {
                            iColumnIndex = scf.Reference[index].StartColumnIndex;
                            iColumnIndex2 = scf.Reference[index].EndColumnIndex;
                            this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref iColumnIndex, ref iColumnIndex2);
                            if (iColumnIndex < 1 || iColumnIndex > SLConstants.ColumnLimit || iColumnIndex2 < 1 || iColumnIndex2 > SLConstants.ColumnLimit)
                            {
                                bToDelete = true;
                                break;
                            }

                            scf.Reference[index] = new SLCellPointRange(
                                scf.Reference[index].StartRowIndex, iColumnIndex,
                                scf.Reference[index].EndRowIndex, iColumnIndex2);
                        }

                        if (bToDelete)
                        {
                            slws.SharedCellFormulas.Remove(scfkeys[i]);
                        }
                    }
                }
                #endregion

                #region Defined names
                if (slwb.DefinedNames.Count > 0)
                {
                    //string sDefinedNameText = string.Empty;
                    //foreach (SLDefinedName d in slwb.DefinedNames)
                    //{
                    //    sDefinedNameText = d.Text;
                    //    sDefinedNameText = AddDeleteCellFormulaDelta(sDefinedNameText, -1, 0, StartColumnIndex, NumberOfColumns);
                    //    sDefinedNameText = AddDeleteDefinedNameRowColumnRangeDelta(sDefinedNameText, false, StartColumnIndex, NumberOfColumns);
                    //    d.Text = sDefinedNameText;
                    //}

                    // not doing anything with this but using it to pass to the function
                    bool bHasError = false;
                    foreach (SLDefinedName d in slwb.DefinedNames)
                    {
                        d.Text = AdjustCellFormulaDelta(d.Text, false, 0, StartColumnIndex, 0, StartColumnIndex + NumberOfColumns, true, true, true, false, StartColumnIndex, StartColumnIndex + NumberOfColumns, out bHasError);
                    }
                }
                #endregion

                #region Sparklines
                if (slws.SparklineGroups.Count > 0)
                {
                    SLSparkline spk;
                    foreach (SLSparklineGroup spkgrp in slws.SparklineGroups)
                    {
                        if (spkgrp.DateAxis && spkgrp.DateWorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref spkgrp.DateStartColumnIndex, ref spkgrp.DateEndColumnIndex);
                        }

                        // starting from the end is important because we might be deleting!
                        for (i = spkgrp.Sparklines.Count - 1; i >= 0; --i)
                        {
                            spk = spkgrp.Sparklines[i];

                            if (spk.LocationColumnIndex >= StartColumnIndex)
                            {
                                iNewIndex = spk.LocationColumnIndex + NumberOfColumns;
                                if (iNewIndex <= SLConstants.ColumnLimit)
                                {
                                    spk.LocationColumnIndex = iNewIndex;
                                }
                                else
                                {
                                    // out of range!
                                    spkgrp.Sparklines.RemoveAt(i);
                                    continue;
                                }
                            }
                            // else the location is before the start column so don't have to do anything

                            // process only if the data source is on the currently selected worksheet
                            if (spk.WorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                this.AddRowColumnIndexDelta(StartColumnIndex, NumberOfColumns, false, ref spk.StartColumnIndex, ref spk.EndColumnIndex);
                            }

                            spkgrp.Sparklines[i] = spk;
                        }
                    }
                }
                #endregion
            }

            return result;
        }

        /// <summary>
        /// Delete one or more columns.
        /// </summary>
        /// <param name="StartColumnName">Columns will deleted from this column, including this column itself.</param>
        /// <param name="NumberOfColumns">Number of columns to delete.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool DeleteColumn(string StartColumnName, int NumberOfColumns)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(StartColumnName);

            return DeleteColumn(iColumnIndex, NumberOfColumns);
        }

        /// <summary>
        /// Delete one or more columns.
        /// </summary>
        /// <param name="StartColumnIndex">Columns will be deleted from this column index, including this column itself.</param>
        /// <param name="NumberOfColumns">Number of columns to delete.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool DeleteColumn(int StartColumnIndex, int NumberOfColumns)
        {
            if (NumberOfColumns < 1) return false;

            bool result = false;
            if (StartColumnIndex >= 1 && StartColumnIndex <= SLConstants.ColumnLimit)
            {
                this.FlattenAllSharedCellFormula();

                result = true;
                int i = 0, iNewIndex = 0;
                int iEndColumnIndex = StartColumnIndex + NumberOfColumns - 1;
                if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;
                // this autocorrects in the case of overshooting the column limit
                int iNumberOfColumns = iEndColumnIndex - StartColumnIndex + 1;

                WorksheetPart wsp;

                int index = 0;
                int iColumnIndex = -1;
                int iColumnIndex2 = -1;

                #region Column properties
                SLColumnProperties cp;
                List<int> listColumnIndex = slws.ColumnProperties.Keys.ToList<int>();
                // this sorting in ascending order is crucial!
                // we're removing data within delete range, and if the data is after the delete range,
                // then the key is modified. Since the data within delete range is removed,
                // there won't be data key collisions.
                listColumnIndex.Sort();

                for (i = 0; i < listColumnIndex.Count; ++i)
                {
                    index = listColumnIndex[i];
                    if (StartColumnIndex <= index && index <= iEndColumnIndex)
                    {
                        slws.ColumnProperties.Remove(index);
                    }
                    else if (index > iEndColumnIndex)
                    {
                        cp = slws.ColumnProperties[index];
                        slws.ColumnProperties.Remove(index);
                        iNewIndex = index - iNumberOfColumns;
                        slws.ColumnProperties[iNewIndex] = cp.Clone();
                    }

                    // the columns before the start column are unaffected by the deleting.
                }
                #endregion

                #region Cell data
                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                // 2 April 2017: Need to sort or the row keys are "random" (depending on
                // how the OS assign memory and order and such)
                // This made the rows "random"
                listRowKeys.Sort();
                List<int> listColumnKeys;

                SLCell c;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    listColumnKeys.Sort();
                    foreach (int colkey in listColumnKeys)
                    {
                        c = slws.CellWarehouse.Cells[rowkey][colkey].Clone();

                        if (StartColumnIndex <= colkey && colkey <= iEndColumnIndex)
                        {
                            slws.CellWarehouse.Remove(rowkey, colkey);
                        }
                        else if (colkey > iEndColumnIndex)
                        {
                            slws.CellWarehouse.Remove(rowkey, colkey);
                            iNewIndex = colkey - iNumberOfColumns;
                            this.ProcessCellFormulaDelta(ref c, false, 0, StartColumnIndex, 0, StartColumnIndex - iNumberOfColumns, true, false, false, StartColumnIndex, iEndColumnIndex);
                            slws.CellWarehouse.SetValue(rowkey, iNewIndex, c);
                        }
                        else
                        {
                            slws.CellWarehouse.SetValue(rowkey, colkey, c);
                        }
                    }
                }

                #region Cell comments
                List<SLCellPoint> listCellRefKeys = slws.Comments.Keys.ToList<SLCellPoint>();
                // this sorting in ascending order is crucial!
                listCellRefKeys.Sort(new SLCellReferencePointComparer());

                SLCellPoint pt;
                SLComment comm;
                for (i = 0; i < listCellRefKeys.Count; ++i)
                {
                    pt = listCellRefKeys[i];
                    comm = slws.Comments[pt];
                    if (StartColumnIndex <= pt.ColumnIndex && pt.ColumnIndex <= iEndColumnIndex)
                    {
                        slws.Comments.Remove(pt);
                    }
                    else if (pt.ColumnIndex > iEndColumnIndex)
                    {
                        slws.Comments.Remove(pt);
                        iNewIndex = pt.ColumnIndex - iNumberOfColumns;
                        slws.Comments[new SLCellPoint(pt.RowIndex, iNewIndex)] = comm;
                    }
                    // no else because there's nothing done
                }
                #endregion

                #endregion

                // Excel doesn't seem to allow inserting/deleting columns in certain cases
                // when 2 (or more) tables overlap each other vertically (meaning one above the other).
                // In particular, when the insert/delete range overlaps an existing column
                // in 2 (or more) tables.
                // The algorithm below works fine, meaning the resulting spreadsheet doesn't
                // have errors, so not sure why Excel doesn't allow it.
                #region Tables
                if (slws.Tables.Count > 0)
                {
                    SLTable t;
                    List<int> listTablesToDelete = new List<int>();
                    for (i = 0; i < slws.Tables.Count; ++i)
                    {
                        t = slws.Tables[i];
                        iColumnIndex = t.StartColumnIndex;
                        iColumnIndex2 = t.EndColumnIndex;
                        if (StartColumnIndex <= iColumnIndex && iColumnIndex2 <= iEndColumnIndex)
                        {
                            // table is completely within delete range, so delete the whole table
                            listTablesToDelete.Add(i);
                            continue;
                        }
                        else
                        {
                            int iTableStartIndex = 0, iNumberOfTableColumnsToDelete = 0;
                            if (StartColumnIndex <= iColumnIndex && iColumnIndex <= iEndColumnIndex && iEndColumnIndex < iColumnIndex2)
                            {
                                // the left part of the table columns are deleted
                                iTableStartIndex = 0;
                                iNumberOfTableColumnsToDelete = iEndColumnIndex - iColumnIndex + 1;
                            }
                            else if (iColumnIndex < StartColumnIndex && iEndColumnIndex < iColumnIndex2)
                            {
                                // the middle part of the table columns are deleted
                                iTableStartIndex = StartColumnIndex - iColumnIndex;
                                iNumberOfTableColumnsToDelete = iEndColumnIndex - StartColumnIndex + 1;
                            }
                            else if (iColumnIndex < StartColumnIndex && StartColumnIndex <= iColumnIndex2 && iColumnIndex2 <= iEndColumnIndex)
                            {
                                // the right part of the table columns are deleted
                                iTableStartIndex = StartColumnIndex - iColumnIndex;
                                iNumberOfTableColumnsToDelete = iColumnIndex2 - StartColumnIndex + 1;
                            }

                            // this assumes that TableColumns only has TableColumn as children
                            // and that the number of table columns corresponds correctly to the
                            // table reference range (because it might be different, but that
                            // means the spreadsheet has an error).
                            // We start from the back because we're deleting so we don't want to
                            // mess up the indices.
                            for (i = iTableStartIndex + iNumberOfTableColumnsToDelete - 1; i >= iTableStartIndex; --i)
                            {
                                t.TableColumns.RemoveAt(i);
                            }

                            this.DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, iNumberOfColumns, ref iColumnIndex, ref iColumnIndex2);
                            if (iColumnIndex != t.StartColumnIndex || iColumnIndex2 != t.EndColumnIndex) t.IsNewTable = true;
                            t.StartColumnIndex = iColumnIndex;
                            t.EndColumnIndex = iColumnIndex2;
                        }

                        if (t.HasAutoFilter)
                        {
                            // if the autofilter range is completely within delete range,
                            // then it's already taken care off above.
                            iColumnIndex = t.AutoFilter.StartColumnIndex;
                            iColumnIndex2 = t.AutoFilter.EndColumnIndex;
                            this.DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, iNumberOfColumns, ref iColumnIndex, ref iColumnIndex2);
                            if (iColumnIndex != t.AutoFilter.StartColumnIndex || iColumnIndex2 != t.AutoFilter.EndColumnIndex) t.IsNewTable = true;
                            t.AutoFilter.StartColumnIndex = iColumnIndex;
                            t.AutoFilter.EndColumnIndex = iColumnIndex2;

                            if (t.AutoFilter.HasSortState)
                            {
                                // if the sort state range is completely within delete range,
                                // then it's already taken care off above.
                                iColumnIndex = t.AutoFilter.SortState.StartColumnIndex;
                                iColumnIndex2 = t.AutoFilter.SortState.EndColumnIndex;
                                this.DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, iNumberOfColumns, ref iColumnIndex, ref iColumnIndex2);
                                if (iColumnIndex != t.AutoFilter.SortState.StartColumnIndex || iColumnIndex2 != t.AutoFilter.SortState.EndColumnIndex) t.IsNewTable = true;
                                t.AutoFilter.SortState.StartColumnIndex = iColumnIndex;
                                t.AutoFilter.SortState.EndColumnIndex = iColumnIndex2;
                            }
                        }

                        if (t.HasSortState)
                        {
                            // if the sort state range is completely within delete range,
                            // then it's already taken care off above.
                            iColumnIndex = t.SortState.StartColumnIndex;
                            iColumnIndex2 = t.SortState.EndColumnIndex;
                            this.DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, iNumberOfColumns, ref iColumnIndex, ref iColumnIndex2);
                            if (iColumnIndex != t.SortState.StartColumnIndex || iColumnIndex2 != t.SortState.EndColumnIndex) t.IsNewTable = true;
                            t.SortState.StartColumnIndex = iColumnIndex;
                            t.SortState.EndColumnIndex = iColumnIndex2;
                        }
                    }

                    if (listTablesToDelete.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
                        {
                            wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                            string sTableRelID = string.Empty;
                            string sTableName = string.Empty;
                            uint iTableID = 0;
                            for (i = listTablesToDelete.Count - 1; i >= 0; --i)
                            {
                                // remove IDs and table names from the spreadsheet unique lists
                                iTableID = slws.Tables[listTablesToDelete[i]].Id;
                                if (slwb.TableIds.Contains(iTableID)) slwb.TableIds.Remove(iTableID);

                                sTableName = slws.Tables[listTablesToDelete[i]].DisplayName;
                                if (slwb.TableNames.Contains(sTableName)) slwb.TableNames.Remove(sTableName);

                                sTableRelID = slws.Tables[listTablesToDelete[i]].RelationshipID;
                                if (sTableRelID.Length > 0)
                                {
                                    wsp.DeletePart(sTableRelID);
                                }
                                slws.Tables.RemoveAt(listTablesToDelete[i]);
                            }
                        }
                    }
                }
                #endregion

                #region Merge cells
                if (slws.MergeCells.Count > 0)
                {
                    SLMergeCell mc;
                    // starting from the end is crucial because we might be removing items
                    for (i = slws.MergeCells.Count - 1; i >= 0; --i)
                    {
                        mc = slws.MergeCells[i];
                        if (StartColumnIndex <= mc.iStartColumnIndex && mc.iEndColumnIndex <= iEndColumnIndex)
                        {
                            // merge cell is completely within delete range
                            slws.MergeCells.RemoveAt(i);
                        }
                        else
                        {
                            this.DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, iNumberOfColumns, ref mc.iStartColumnIndex, ref mc.iEndColumnIndex);
                            slws.MergeCells[i] = mc;
                        }
                    }
                }
                #endregion

                #region Hyperlinks
                if (slws.Hyperlinks.Count > 0)
                {
                    List<SLCellPoint> deletepoints = new List<SLCellPoint>();
                    SLHyperlink hl;
                    for (i = slws.Hyperlinks.Count - 1; i >= 0; --i)
                    {
                        hl = slws.Hyperlinks[i];
                        if (StartColumnIndex <= hl.Reference.StartColumnIndex && hl.Reference.EndColumnIndex <= iEndColumnIndex)
                        {
                            // hyperlink is completely within delete range
                            // We use a list to take of the points because we also need to remove any existing
                            // hyperlink relationship. It's easier to just use the remove hyperlink function.
                            deletepoints.Add(new SLCellPoint(hl.Reference.StartRowIndex, hl.Reference.StartColumnIndex));
                        }
                        else
                        {
                            iColumnIndex = hl.Reference.StartColumnIndex;
                            iColumnIndex2 = hl.Reference.EndColumnIndex;
                            this.DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, iNumberOfColumns, ref iColumnIndex, ref iColumnIndex2);
                            hl.Reference = new SLCellPointRange(hl.Reference.StartRowIndex, iColumnIndex, hl.Reference.EndRowIndex, iColumnIndex2);
                            slws.Hyperlinks[i] = hl.Clone();
                        }
                    }

                    foreach (SLCellPoint hyperlinkpt in deletepoints)
                    {
                        this.RemoveHyperlink(hyperlinkpt.RowIndex, hyperlinkpt.ColumnIndex);
                    }
                }
                #endregion

                #region Drawings
                if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
                {
                    wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                    if (wsp.DrawingsPart != null)
                    {
                        bool bFound = false;
                        bool bToFindCumulative = false;
                        Xdr.TwoCellAnchor tcaNew;
                        Xdr.OneCellAnchor ocaNew;
                        Xdr.EditAsValues vEditAs = Xdr.EditAsValues.Absolute;
                        List<OpenXmlElement> listoxe = new List<OpenXmlElement>();
                        int iIndex = 0;
                        int iIndex2 = 0;

                        long lDefaultWidth = slws.SheetFormatProperties.DefaultColumnWidthInEMU;
                        Dictionary<int, long> dictColumnWidth = new Dictionary<int, long>();
                        List<int> listindex = slws.ColumnProperties.Keys.ToList<int>();
                        foreach (int colindex in listindex)
                        {
                            cp = slws.ColumnProperties[colindex];
                            dictColumnWidth[colindex] = cp.WidthInEMU;
                        }

                        Dictionary<int, long> dictCumulativeLength = new Dictionary<int, long>();

                        long lStartDelete = 0, lEndDelete = 0;
                        long lAnchorStartDelete = 0, lAnchorEndDelete = 0;

                        long lCumulative = 0;
                        for (iIndex = 1; iIndex <= iEndColumnIndex; ++iIndex)
                        {
                            if (dictColumnWidth.ContainsKey(iIndex))
                            {
                                lCumulative += dictColumnWidth[iIndex];
                            }
                            else
                            {
                                lCumulative += lDefaultWidth;
                            }

                            dictCumulativeLength[iIndex] = lCumulative;
                        }

                        // we're getting the start and end "lines" of the delete range in terms of EMUs.
                        // The start line is at the left of the delete range, which is the right of the
                        // previous column. The end line is at the right of the delete range, which
                        // is exactly the current column.

                        if (StartColumnIndex > 1) lStartDelete = dictCumulativeLength[StartColumnIndex - 1];
                        else lStartDelete = 0;

                        lEndDelete = dictCumulativeLength[iEndColumnIndex];

                        DrawingsPart dp = wsp.DrawingsPart;
                        foreach (OpenXmlElement oxe in dp.WorksheetDrawing.ChildElements)
                        {
                            if (oxe is Xdr.TwoCellAnchor)
                            {
                                tcaNew = (Xdr.TwoCellAnchor)oxe.CloneNode(true);

                                if (tcaNew.EditAs == null) vEditAs = Xdr.EditAsValues.TwoCell;
                                else vEditAs = tcaNew.EditAs.Value;

                                if (vEditAs != Xdr.EditAsValues.Absolute
                                    && tcaNew.FromMarker != null && tcaNew.FromMarker.ColumnId != null
                                    && tcaNew.ToMarker != null && tcaNew.ToMarker.ColumnId != null)
                                {
                                    iIndex = Convert.ToInt32(tcaNew.FromMarker.ColumnId.Text);
                                    iIndex2 = Convert.ToInt32(tcaNew.ToMarker.ColumnId.Text);

                                    if (iIndex2 + 1 > dictCumulativeLength.Count)
                                    {
                                        lCumulative = dictCumulativeLength[dictCumulativeLength.Count];
                                        for (i = dictCumulativeLength.Count + 1; i <= iIndex2 + 1; ++i)
                                        {
                                            if (dictColumnWidth.ContainsKey(i))
                                            {
                                                lCumulative += dictColumnWidth[i];
                                            }
                                            else
                                            {
                                                lCumulative += lDefaultWidth;
                                            }
                                            dictCumulativeLength[i] = lCumulative;
                                        }
                                    }

                                    // the index is 0-based while the spreadsheet index is 1-based.

                                    if (iIndex2 > 0) lAnchorEndDelete = dictCumulativeLength[iIndex2];
                                    else lAnchorEndDelete = 0;

                                    if (tcaNew.ToMarker.ColumnOffset != null)
                                    {
                                        lAnchorEndDelete += Convert.ToInt64(tcaNew.ToMarker.ColumnOffset.Text);
                                    }

                                    if (lAnchorEndDelete > dictCumulativeLength[dictCumulativeLength.Count])
                                    {
                                        lCumulative = dictCumulativeLength[dictCumulativeLength.Count];
                                        while (lAnchorEndDelete > lCumulative)
                                        {
                                            i = dictCumulativeLength.Count + 1;
                                            if (dictColumnWidth.ContainsKey(i))
                                            {
                                                lCumulative += dictColumnWidth[i];
                                            }
                                            else
                                            {
                                                lCumulative += lDefaultWidth;
                                            }
                                            dictCumulativeLength[i] = lCumulative;
                                        }
                                    }

                                    // assume the ToMarker index is after the FromMarker, and that after
                                    // taking into account the offset, it's still before the ToMarker,
                                    // so we don't do any more checks

                                    if (iIndex > 0) lAnchorStartDelete = dictCumulativeLength[iIndex];
                                    else lAnchorStartDelete = 0;

                                    if (tcaNew.FromMarker.ColumnOffset != null)
                                    {
                                        lAnchorStartDelete += Convert.ToInt64(tcaNew.FromMarker.ColumnOffset.Text);
                                    }

                                    bToFindCumulative = false;
                                    if (vEditAs == Xdr.EditAsValues.TwoCell)
                                    {
                                        if (lStartDelete <= lAnchorStartDelete && lAnchorEndDelete <= lEndDelete)
                                        {
                                            // completely within delete range
                                            // Not sure how to handle any hyperlinks tied to drawing or physical media files.
                                            // So will just make the drawing very very small. Say 1 EMU.
                                            // That's practically invisible.
                                            tcaNew.FromMarker.ColumnOffset = new Xdr.ColumnOffset("0");
                                            tcaNew.ToMarker.ColumnId = new Xdr.ColumnId(tcaNew.FromMarker.ColumnId.Text);
                                            tcaNew.ToMarker.ColumnOffset = new Xdr.ColumnOffset("1");
                                            bFound = true;
                                            bToFindCumulative = false;
                                        }
                                        else if (lEndDelete <= lAnchorStartDelete)
                                        {
                                            // delete range is before drawing
                                            lCumulative = lEndDelete - lStartDelete;
                                            lAnchorEndDelete -= lCumulative;
                                            lAnchorStartDelete -= lCumulative;

                                            bToFindCumulative = true;
                                        }
                                        else if (lStartDelete <= lAnchorStartDelete && lAnchorStartDelete <= lEndDelete && lEndDelete <= lAnchorEndDelete)
                                        {
                                            // left part is within delete range
                                            lCumulative = lEndDelete - lAnchorStartDelete;
                                            lAnchorEndDelete -= lCumulative;

                                            // put the drawing flush at the start of delete range
                                            lCumulative = lAnchorEndDelete - lAnchorStartDelete;
                                            lAnchorStartDelete = lStartDelete;
                                            lAnchorEndDelete = lAnchorStartDelete + lCumulative;

                                            bToFindCumulative = true;
                                        }
                                        else if (lAnchorStartDelete <= lStartDelete && lEndDelete <= lAnchorEndDelete)
                                        {
                                            // delete range is within drawing
                                            lCumulative = lEndDelete - lStartDelete;
                                            lAnchorEndDelete -= lCumulative;

                                            bToFindCumulative = true;
                                        }
                                        else if (lAnchorStartDelete <= lStartDelete && lStartDelete <= lAnchorEndDelete && lAnchorEndDelete <= lEndDelete)
                                        {
                                            // right part is within delete range
                                            lAnchorEndDelete = lStartDelete;

                                            bToFindCumulative = true;
                                        }
                                        // else the delete range is beyond the drawing, so nothing needs to be done
                                    }
                                    else if (vEditAs == Xdr.EditAsValues.OneCell)
                                    {
                                        if (lStartDelete <= lAnchorStartDelete)
                                        {
                                            // hold temporarily
                                            lCumulative = lAnchorStartDelete;

                                            lAnchorStartDelete -= (lEndDelete - lStartDelete);
                                            if (lAnchorStartDelete < lStartDelete) lAnchorStartDelete = lStartDelete;

                                            lAnchorEndDelete -= (lCumulative - lAnchorStartDelete);
                                            // the above algorithm makes it such that we don't overshoot the start delete
                                            // range. Normally, we'd only have the anchored start (FromMarker), but
                                            // TwoCellAnchor's have the ToMarker too, which is a pain in the Adam's apple.

                                            bToFindCumulative = true;
                                        }
                                    }

                                    if (bToFindCumulative)
                                    {
                                        iIndex = -1;
                                        iIndex2 = -1;
                                        for (i = dictCumulativeLength.Count; i >= 0; --i)
                                        {
                                            if (i == 0)
                                            {
                                                if (iIndex < 0)
                                                {
                                                    iIndex = 0;
                                                    tcaNew.FromMarker.ColumnId = new Xdr.ColumnId(iIndex.ToString(CultureInfo.InvariantCulture));
                                                    tcaNew.FromMarker.ColumnOffset = new Xdr.ColumnOffset(lAnchorStartDelete.ToString(CultureInfo.InvariantCulture));
                                                }

                                                if (iIndex2 < 0)
                                                {
                                                    iIndex2 = 0;
                                                    tcaNew.ToMarker.ColumnId = new Xdr.ColumnId(iIndex2.ToString(CultureInfo.InvariantCulture));
                                                    tcaNew.ToMarker.ColumnOffset = new Xdr.ColumnOffset(lAnchorEndDelete.ToString(CultureInfo.InvariantCulture));
                                                }
                                            }

                                            if (iIndex < 0 && lAnchorStartDelete >= dictCumulativeLength[i])
                                            {
                                                iIndex = i;
                                                tcaNew.FromMarker.ColumnId = new Xdr.ColumnId(iIndex.ToString(CultureInfo.InvariantCulture));
                                                tcaNew.FromMarker.ColumnOffset = new Xdr.ColumnOffset((lAnchorStartDelete - dictCumulativeLength[i]).ToString(CultureInfo.InvariantCulture));
                                            }
                                            if (iIndex2 < 0 && lAnchorEndDelete >= dictCumulativeLength[i])
                                            {
                                                iIndex2 = i;
                                                tcaNew.ToMarker.ColumnId = new Xdr.ColumnId(iIndex2.ToString(CultureInfo.InvariantCulture));
                                                tcaNew.ToMarker.ColumnOffset = new Xdr.ColumnOffset((lAnchorEndDelete - dictCumulativeLength[i]).ToString(CultureInfo.InvariantCulture));
                                            }

                                            if (iIndex >= 0 && iIndex2 >= 0) break;
                                        }

                                        if (iIndex >= 0 && iIndex2 >= 0)
                                        {
                                            bFound = true;
                                        }
                                    }
                                }

                                // Need to do for Transform for the child elements?

                                listoxe.Add(tcaNew.CloneNode(true));
                            }
                            else if (oxe is Xdr.OneCellAnchor)
                            {
                                ocaNew = (Xdr.OneCellAnchor)oxe.CloneNode(true);
                                if (ocaNew.FromMarker != null && ocaNew.FromMarker.ColumnId != null)
                                {
                                    iIndex = Convert.ToInt32(ocaNew.FromMarker.ColumnId.Text);
                                    // the index is 0-based while the spreadsheet index is 1-based.
                                    if ((iIndex + 1) >= StartColumnIndex)
                                    {
                                        iIndex -= iNumberOfColumns;
                                        // Excel seems to stop drawings from overshooting
                                        if ((iIndex + 1) < StartColumnIndex)
                                        {
                                            iIndex = StartColumnIndex - 1;
                                            // to put drawing flush with row
                                            if (ocaNew.Extent != null && ocaNew.Extent.Cx != null)
                                            {
                                                ocaNew.Extent.Cx = 0;
                                            }
                                        }
                                        bFound = true;
                                        ocaNew.FromMarker.ColumnId.Text = iIndex.ToString(CultureInfo.InvariantCulture);
                                    }
                                }

                                // Need to do for Transform for the child elements?

                                listoxe.Add(ocaNew.CloneNode(true));
                            }
                            else
                            {
                                listoxe.Add(oxe.CloneNode(true));
                            }
                        }

                        bFound = true;
                        if (bFound)
                        {
                            wsp.DrawingsPart.WorksheetDrawing.RemoveAllChildren();
                            foreach (OpenXmlElement oxe in listoxe)
                            {
                                wsp.DrawingsPart.WorksheetDrawing.Append(oxe.CloneNode(true));
                            }
                            wsp.DrawingsPart.WorksheetDrawing.Save();
                        }
                    }
                }
                #endregion

                // TODO: chart series references

                #region Calculation chain
                if (slwb.CalculationCells.Count > 0)
                {
                    // I don't know enough to fiddle with calculation chains. So I'm going to ignore it.
                    slwb.CalculationCells.Clear();

                    //List<int> listToDelete = new List<int>();
                    //for (i = 0; i < slwb.CalculationCells.Count; ++i)
                    //{
                    //    if (slwb.CalculationCells[i].SheetId == giSelectedWorksheetID)
                    //    {
                    //        if (StartColumnIndex <= slwb.CalculationCells[i].ColumnIndex && slwb.CalculationCells[i].ColumnIndex <= iEndColumnIndex)
                    //        {
                    //            listToDelete.Add(i);
                    //        }
                    //        else if (iEndColumnIndex < slwb.CalculationCells[i].ColumnIndex)
                    //        {
                    //            slwb.CalculationCells[i].ColumnIndex -= iNumberOfColumns;
                    //        }
                    //    }
                    //}

                    //// start from the back because we're deleting elements and we don't want
                    //// the indices to get messed up.
                    //for (i = listToDelete.Count - 1; i >= 0; --i)
                    //{
                    //    slwb.CalculationCells.RemoveAt(listToDelete[i]);
                    //}
                }
                #endregion

                #region Shared formulas
                if (slws.SharedCellFormulas.Count > 0)
                {
                    SLSharedCellFormula scf;
                    bool bToDelete;
                    List<uint> scfkeys = slws.SharedCellFormulas.Keys.ToList<uint>();
                    for (i = scfkeys.Count - 1; i >= 0; --i)
                    {
                        scf = slws.SharedCellFormulas[scfkeys[i]];
                        bToDelete = false;
                        // start from the end because we might be deleting
                        for (index = scf.Reference.Count - 1; index >= 0; --index)
                        {
                            iColumnIndex = scf.Reference[index].StartColumnIndex;
                            iColumnIndex2 = scf.Reference[index].EndColumnIndex;
                            this.DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, iNumberOfColumns, ref iColumnIndex, ref iColumnIndex2);
                            if (iColumnIndex < 1 || iColumnIndex > SLConstants.ColumnLimit || iColumnIndex2 < 1 || iColumnIndex2 > SLConstants.ColumnLimit)
                            {
                                bToDelete = true;
                                break;
                            }

                            scf.Reference[index] = new SLCellPointRange(
                                scf.Reference[index].StartRowIndex, iColumnIndex,
                                scf.Reference[index].EndRowIndex, iColumnIndex2);
                        }

                        if (bToDelete)
                        {
                            slws.SharedCellFormulas.Remove(scfkeys[i]);
                        }
                    }
                }
                #endregion

                #region Defined names
                if (slwb.DefinedNames.Count > 0)
                {
                    //string sDefinedNameText = string.Empty;
                    //foreach (SLDefinedName d in slwb.DefinedNames)
                    //{
                    //    sDefinedNameText = d.Text;
                    //    sDefinedNameText = AddDeleteCellFormulaDelta(sDefinedNameText, -1, 0, StartColumnIndex, -NumberOfColumns);
                    //    sDefinedNameText = AddDeleteDefinedNameRowColumnRangeDelta(sDefinedNameText, false, StartColumnIndex, -NumberOfColumns);
                    //    d.Text = sDefinedNameText;
                    //}

                    // not doing anything with this but using it to pass to the function
                    bool bHasError = false;
                    foreach (SLDefinedName d in slwb.DefinedNames)
                    {
                        d.Text = AdjustCellFormulaDelta(d.Text, false, 0, StartColumnIndex, 0, StartColumnIndex - NumberOfColumns, true, true, false, false, StartColumnIndex, iEndColumnIndex, out bHasError);
                    }
                }
                #endregion

                #region Sparklines
                if (slws.SparklineGroups.Count > 0)
                {
                    SLSparkline spk;
                    foreach (SLSparklineGroup spkgrp in slws.SparklineGroups)
                    {
                        if (spkgrp.DateAxis && spkgrp.DateWorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (StartColumnIndex <= spkgrp.DateStartColumnIndex && spkgrp.DateEndColumnIndex <= iEndColumnIndex)
                            {
                                // the whole date range is completely within delete range
                                spkgrp.DateAxis = false;
                            }
                            else
                            {
                                this.DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, iNumberOfColumns, ref spkgrp.DateStartColumnIndex, ref spkgrp.DateEndColumnIndex);
                            }
                        }

                        // starting from the end is important because we might be deleting!
                        for (i = spkgrp.Sparklines.Count - 1; i >= 0; --i)
                        {
                            spk = spkgrp.Sparklines[i];

                            if (StartColumnIndex <= spk.LocationColumnIndex && spk.LocationColumnIndex <= iEndColumnIndex)
                            {
                                spkgrp.Sparklines.RemoveAt(i);
                                continue;
                            }
                            else if (spk.LocationColumnIndex > iEndColumnIndex)
                            {
                                iNewIndex = spk.LocationColumnIndex - iNumberOfColumns;
                                spk.LocationColumnIndex = iNewIndex;
                            }
                            // no else because there's nothing done

                            // process only if the data source is on the currently selected worksheet
                            if (spk.WorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                if (StartColumnIndex <= spk.StartColumnIndex && spk.EndColumnIndex <= iEndColumnIndex)
                                {
                                    // the data source is completely within delete range
                                    // Excel 2010 keeps the WorksheetExtension, but I'm gonna just delete the whole thing.
                                    spkgrp.Sparklines.RemoveAt(i);
                                    continue;
                                }
                                else
                                {
                                    this.DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, iNumberOfColumns, ref spk.StartColumnIndex, ref spk.EndColumnIndex);
                                }
                            }

                            spkgrp.Sparklines[i] = spk;
                        }
                    }
                }
                #endregion
            }

            return result;
        }

        /// <summary>
        /// Clear all cell content within specified columns. If the top-left cell of a merged cell is within specified columns, the merged cell content is also cleared.
        /// </summary>
        /// <param name="StartColumnName">The column name of the start column.</param>
        /// <param name="EndColumnName">The column name of the end column.</param>
        /// <returns>True if content has been cleared. False otherwise. If there are no content within specified rows, false is also returned.</returns>
        public bool ClearColumnContent(string StartColumnName, string EndColumnName)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            return ClearColumnContent(iStartColumnIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Clear all cell content within specified columns. If the top-left cell of a merged cell is within specified columns, the merged cell content is also cleared.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        /// <returns>True if content has been cleared. False otherwise. If there are no content within specified rows, false is also returned.</returns>
        public bool ClearColumnContent(int StartColumnIndex, int EndColumnIndex)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            bool result = false;
            int i = 0;
            for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
            {
                if (slws.ColumnProperties.ContainsKey(i))
                {
                    if (slws.ColumnProperties[i].IsEmpty)
                    {
                        slws.ColumnProperties.Remove(i);
                        result = true;
                    }
                }
            }

            List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
            List<int> listColumnKeys;
            foreach (int rowkey in listRowKeys)
            {
                listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                foreach (int colkey in listColumnKeys)
                {
                    if (iStartColumnIndex <= colkey && colkey <= iEndColumnIndex)
                    {
                        this.ClearCellContentData(rowkey, colkey);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Delta is >= 0
        /// </summary>
        /// <param name="GivenStartIndex"></param>
        /// <param name="Delta">Delta is >= 0</param>
        /// <param name="IsRow"></param>
        /// <param name="CurrentStartIndex"></param>
        /// <param name="CurrentEndIndex"></param>
        internal void AddRowColumnIndexDelta(int GivenStartIndex, int Delta, bool IsRow, ref int CurrentStartIndex, ref int CurrentEndIndex)
        {
            if (CurrentStartIndex >= GivenStartIndex)
            {
                CurrentStartIndex += Delta;
                if (IsRow)
                {
                    if (CurrentStartIndex > SLConstants.RowLimit) CurrentStartIndex = SLConstants.RowLimit;
                }
                else
                {
                    if (CurrentStartIndex > SLConstants.ColumnLimit) CurrentStartIndex = SLConstants.ColumnLimit;
                }
            }

            if (CurrentEndIndex >= GivenStartIndex)
            {
                CurrentEndIndex += Delta;
                if (IsRow)
                {
                    if (CurrentEndIndex > SLConstants.RowLimit) CurrentEndIndex = SLConstants.RowLimit;
                }
                else
                {
                    if (CurrentEndIndex > SLConstants.ColumnLimit) CurrentEndIndex = SLConstants.ColumnLimit;
                }
            }
        }

        /// <summary>
        /// Delta is >= 0
        /// </summary>
        /// <param name="GivenStartIndex"></param>
        /// <param name="GivenEndIndex"></param>
        /// <param name="Delta">Delta is >= 0</param>
        /// <param name="CurrentStartIndex"></param>
        /// <param name="CurrentEndIndex"></param>
        internal void DeleteRowColumnIndexDelta(int GivenStartIndex, int GivenEndIndex, int Delta, ref int CurrentStartIndex, ref int CurrentEndIndex)
        {
            // the case where the current range is completely within the delete range
            // should already be handled by the calling function.

            if (GivenEndIndex < CurrentStartIndex)
            {
                // current range is completely below/right-of delete range
                CurrentStartIndex -= Delta;
                CurrentEndIndex -= Delta;
            }
            else if ((GivenStartIndex <= CurrentStartIndex && CurrentStartIndex <= GivenEndIndex) && GivenEndIndex < CurrentEndIndex)
            {
                // top/left part of current range is within delete range
                CurrentStartIndex = GivenEndIndex + 1;
                CurrentStartIndex -= Delta;
                CurrentEndIndex -= Delta;
            }
            else if (CurrentStartIndex < GivenStartIndex && GivenEndIndex < CurrentEndIndex)
            {
                // current range strictly covers the delete range
                CurrentEndIndex -= Delta;
            }
            else if (CurrentStartIndex < GivenStartIndex && (GivenStartIndex <= CurrentEndIndex && CurrentEndIndex <= GivenEndIndex))
            {
                // bottom/right part of current range is within delete range
                // That part is gone, so move the end index to 1 level before the given start index
                CurrentEndIndex = GivenStartIndex - 1;
            }

            // else the delete range is complete below/right-of the current range
            // so don't have to do anything
        }

        /// <summary>
        /// This returns a list of index with pixel lengths. Depending on the type,
        /// the pixel length is for row heights or column widths
        /// </summary>
        /// <param name="IsRow"></param>
        /// <param name="StartIndex"></param>
        /// <param name="EndIndex"></param>
        /// <param name="MaxPixelLength"></param>
        /// <returns></returns>
        internal Dictionary<int, int> AutoFitRowColumn(bool IsRow, int StartIndex, int EndIndex, int MaxPixelLength)
        {
            int i;
            Dictionary<int, int> pixellength = new Dictionary<int, int>();
            // initialise all to zero first. This also ensures the existence of a dictionary entry.
            for (i = StartIndex; i <= EndIndex; ++i)
            {
                pixellength[i] = 0;
            }

            List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
            List<int> listColumnKeys;
            Dictionary<int, HashSet<int>> multikeys = new Dictionary<int, HashSet<int>>();
            foreach (int rowkey in listRowKeys)
            {
                multikeys.Add(rowkey, new HashSet<int>());
                listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                foreach (int colkey in listColumnKeys)
                {
                    multikeys[rowkey].Add(colkey);
                }
            }

            SLCell c;
            string sAutoFitSharedStringCacheKey;
            string sAutoFitCacheKey;
            double fCellValue;
            SLRstType rst;
            Text txt;
            Run run;
            FontSchemeValues vFontScheme;
            int index;
            SLStyle style = new SLStyle();
            int iStyleIndex;
            string sFontName;
            double fFontSize;
            bool bBold;
            bool bItalic;
            bool bStrike;
            bool bUnderline;
            System.Drawing.FontStyle drawstyle;
            System.Drawing.Font ftUsable;
            string sFormatCode;
            string sDotNetFormatCode;
            int iTextRotation;
            System.Drawing.SizeF szf;
            string sText;
            float fWidth;
            float fHeight;
            int iPointIndex;
            bool bSkipAdjustment;

            SLCellPoint ptCheck;
            // remove points that are part of merged cells
            // Merged cells don't factor into autofitting.
            // Start from end because we will be deleting points (that's why we're reversing).
            if (slws.MergeCells.Count > 0)
            {
                listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                listRowKeys.Sort();
                listRowKeys.Reverse();
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    listColumnKeys.Sort();
                    listColumnKeys.Reverse();
                    foreach (int colkey in listColumnKeys)
                    {
                        foreach (SLMergeCell mc in slws.MergeCells)
                        {
                            if (mc.StartRowIndex <= rowkey && rowkey <= mc.EndRowIndex
                                && mc.StartColumnIndex <= colkey && colkey <= mc.EndColumnIndex)
                            {
                                if (multikeys.ContainsKey(rowkey) && multikeys[rowkey].Contains(colkey))
                                {
                                    multikeys[rowkey].Remove(colkey);
                                }
                                break;
                            }
                        }
                    }
                }
            }

            HashSet<SLCellPoint> hsFilter = new HashSet<SLCellPoint>();
            if (slws.HasAutoFilter)
            {
                for (i = slws.AutoFilter.StartColumnIndex; i <= slws.AutoFilter.EndColumnIndex; ++i)
                {
                    hsFilter.Add(new SLCellPoint(slws.AutoFilter.StartRowIndex, i));
                }
            }

            if (slws.Tables.Count > 0)
            {
                foreach (SLTable t in slws.Tables)
                {
                    if (t.HasAutoFilter)
                    {
                        for (i = t.AutoFilter.StartColumnIndex; i <= t.AutoFilter.EndColumnIndex; ++i)
                        {
                            ptCheck = new SLCellPoint(t.AutoFilter.StartRowIndex, i);
                            if (!hsFilter.Contains(ptCheck))
                            {
                                hsFilter.Add(ptCheck);
                            }
                        }
                    }
                }
            }

            // Excel seems to stop the maximum column pixel width at 2300 pixels (at least at 120 DPI).
            // We need a bitmap of sufficient size because we're rendering the text and measuring it.
            // 4096 pixels wide should cover the 2300 pixel thing. Note that this is also wider than
            // typical screen monitors.
            // 2048 pixels high should also cover most screen monitors' vertical height.
            // If your text fills up the entire height of your screen, I would say your font size is
            // too large...
            // If you're doing this in some distant future where you can do spreadsheets on the
            // freaking wall with screens of cosmic proportions, feel free to increase the dimensions.
            using (System.Drawing.Bitmap bm = new System.Drawing.Bitmap(4096, 2048))
            {
                using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(bm))
                {
                    listRowKeys = multikeys.Keys.ToList<int>();
                    foreach (int rowkey in listRowKeys)
                    {
                        listColumnKeys = multikeys[rowkey].ToList<int>();
                        foreach (int colkey in listColumnKeys)
                        {
                            if (IsRow) iPointIndex = rowkey;
                            else iPointIndex = colkey;

                            if (StartIndex <= iPointIndex && iPointIndex <= EndIndex)
                            {
                                c = slws.CellWarehouse.Cells[rowkey][colkey].Clone();

                                iStyleIndex = (int)c.StyleIndex;
                                // assume if the font cache contains the style index,
                                // the other caches also have it because all the caches are set
                                // at the same time and they are all internal (as in not public).
                                if (dictAutoFitFontCache.ContainsKey(iStyleIndex))
                                {
                                    ftUsable = dictAutoFitFontCache[iStyleIndex];
                                    sDotNetFormatCode = dictAutoFitFormatCodeCache[iStyleIndex];
                                    sFormatCode = sDotNetFormatCode;
                                    iTextRotation = dictAutoFitTextRotationCache[iStyleIndex];
                                }
                                else
                                {
                                    style = new SLStyle();
                                    style.FromHash(listStyle[iStyleIndex]);

                                    #region Get style stuff
                                    sFontName = SimpleTheme.MinorLatinFont;
                                    fFontSize = SLConstants.DefaultFontSize;
                                    bBold = false;
                                    bItalic = false;
                                    bStrike = false;
                                    bUnderline = false;
                                    drawstyle = System.Drawing.FontStyle.Regular;
                                    if (style.HasFont)
                                    {
                                        if (style.fontReal.HasFontScheme)
                                        {
                                            if (style.fontReal.FontScheme == FontSchemeValues.Major) sFontName = SimpleTheme.MajorLatinFont;
                                            else if (style.fontReal.FontScheme == FontSchemeValues.Minor) sFontName = SimpleTheme.MinorLatinFont;
                                            else if (style.fontReal.FontName != null && style.fontReal.FontName.Length > 0) sFontName = style.fontReal.FontName;
                                        }
                                        else
                                        {
                                            if (style.fontReal.FontName != null && style.fontReal.FontName.Length > 0) sFontName = style.fontReal.FontName;
                                        }

                                        if (style.fontReal.FontSize != null) fFontSize = style.fontReal.FontSize.Value;
                                        if (style.fontReal.Bold != null && style.fontReal.Bold.Value) bBold = true;
                                        if (style.fontReal.Italic != null && style.fontReal.Italic.Value) bItalic = true;
                                        if (style.fontReal.Strike != null && style.fontReal.Strike.Value) bStrike = true;
                                        if (style.fontReal.HasUnderline) bUnderline = true;
                                    }

                                    if (bBold) drawstyle |= System.Drawing.FontStyle.Bold;
                                    if (bItalic) drawstyle |= System.Drawing.FontStyle.Italic;
                                    if (bStrike) drawstyle |= System.Drawing.FontStyle.Strikeout;
                                    if (bUnderline) drawstyle |= System.Drawing.FontStyle.Underline;

                                    ftUsable = SLTool.GetUsableNormalFont(sFontName, fFontSize, drawstyle, gbThrowExceptionsIfAny);
                                    sFormatCode = style.FormatCode;
                                    sDotNetFormatCode = SLTool.ToDotNetFormatCode(sFormatCode);
                                    if (style.HasAlignment && style.alignReal.TextRotation != null)
                                    {
                                        iTextRotation = style.alignReal.TextRotation.Value;
                                    }
                                    else
                                    {
                                        iTextRotation = 0;
                                    }

                                    #endregion

                                    dictAutoFitFontCache[iStyleIndex] = (System.Drawing.Font)ftUsable.Clone();
                                    dictAutoFitFormatCodeCache[iStyleIndex] = sDotNetFormatCode;
                                    dictAutoFitTextRotationCache[iStyleIndex] = iTextRotation;
                                }

                                sText = string.Empty;

                                fWidth = 0;
                                fHeight = 0;
                                // must set empty first! Used for checking if shared string and if should set into cache.
                                sAutoFitSharedStringCacheKey = string.Empty;
                                bSkipAdjustment = false;

                                if (c.DataType == CellValues.SharedString)
                                {
                                    index = Convert.ToInt32(c.NumericValue);

                                    sAutoFitSharedStringCacheKey = string.Format("{0}{1}{2}",
                                        index.ToString(CultureInfo.InvariantCulture),
                                        SLConstants.AutoFitCacheSeparator,
                                        c.StyleIndex.ToString(CultureInfo.InvariantCulture));
                                    if (dictAutoFitSharedStringCache.ContainsKey(sAutoFitSharedStringCacheKey))
                                    {
                                        fHeight = dictAutoFitSharedStringCache[sAutoFitSharedStringCacheKey].Height;
                                        fWidth = dictAutoFitSharedStringCache[sAutoFitSharedStringCacheKey].Width;
                                        bSkipAdjustment = true;
                                    }
                                    else if (index >= 0 && index < listSharedString.Count)
                                    {
                                        rst = new SLRstType();
                                        rst.FromHash(listSharedString[index]);

                                        if (rst.istrReal.ChildElements.Count == 1 && rst.istrReal.Text != null)
                                        {
                                            sText = rst.istrReal.Text.Text.TrimEnd();
                                            sAutoFitCacheKey = string.Format("{0}{1}{2}", sText, SLConstants.AutoFitCacheSeparator, iStyleIndex.ToString(CultureInfo.InvariantCulture));

                                            if (dictAutoFitTextCache.ContainsKey(sAutoFitCacheKey))
                                            {
                                                szf = dictAutoFitTextCache[sAutoFitCacheKey];
                                                fHeight = szf.Height;
                                                fWidth = szf.Width;
                                            }
                                            else
                                            {
                                                szf = SLTool.MeasureText(bm, g, sText, ftUsable);
                                                fHeight = szf.Height;
                                                fWidth = szf.Width;
                                                dictAutoFitTextCache[sAutoFitCacheKey] = new System.Drawing.SizeF(fWidth, fHeight);
                                            }
                                        }
                                        else
                                        {
                                            i = 0;
                                            foreach (var child in rst.istrReal.ChildElements.Reverse())
                                            {
                                                if (child is Text || child is Run)
                                                {
                                                    if (child is Text)
                                                    {
                                                        txt = (Text)child;
                                                        sText = txt.Text;

                                                        // the last element has the trailing spaces ignored. Hence the Reverse() above.
                                                        if (i == 0) sText = sText.TrimEnd();

                                                        szf = SLTool.MeasureText(bm, g, sText, ftUsable);
                                                        if (szf.Height > fHeight) fHeight = szf.Height;
                                                        fWidth += szf.Width;
                                                    }
                                                    else if (child is Run)
                                                    {
                                                        sText = string.Empty;
                                                        sFontName = (ftUsable.Name != null && ftUsable.Name.Length > 0) ? ftUsable.Name : SimpleTheme.MinorLatinFont;
                                                        fFontSize = ftUsable.SizeInPoints;
                                                        bBold = ((ftUsable.Style & System.Drawing.FontStyle.Bold) > 0) ? true : false;
                                                        bItalic = ((ftUsable.Style & System.Drawing.FontStyle.Italic) > 0) ? true : false;
                                                        bStrike = ((ftUsable.Style & System.Drawing.FontStyle.Strikeout) > 0) ? true : false;
                                                        bUnderline = ((ftUsable.Style & System.Drawing.FontStyle.Underline) > 0) ? true : false;
                                                        drawstyle = System.Drawing.FontStyle.Regular;

                                                        run = (Run)child;
                                                        sText = run.Text.Text;
                                                        vFontScheme = FontSchemeValues.None;
                                                        #region Run properties
                                                        if (run.RunProperties != null)
                                                        {
                                                            foreach (var grandchild in run.RunProperties.ChildElements)
                                                            {
                                                                if (grandchild is RunFont)
                                                                {
                                                                    sFontName = ((RunFont)grandchild).Val;
                                                                }
                                                                else if (grandchild is FontSize)
                                                                {
                                                                    fFontSize = ((FontSize)grandchild).Val;
                                                                }
                                                                else if (grandchild is Bold)
                                                                {
                                                                    Bold b = (Bold)grandchild;
                                                                    if (b.Val == null) bBold = true;
                                                                    else bBold = b.Val.Value;
                                                                }
                                                                else if (grandchild is Italic)
                                                                {
                                                                    Italic itlc = (Italic)grandchild;
                                                                    if (itlc.Val == null) bItalic = true;
                                                                    else bItalic = itlc.Val.Value;
                                                                }
                                                                else if (grandchild is Strike)
                                                                {
                                                                    Strike strk = (Strike)grandchild;
                                                                    if (strk.Val == null) bStrike = true;
                                                                    else bStrike = strk.Val.Value;
                                                                }
                                                                else if (grandchild is Underline)
                                                                {
                                                                    Underline und = (Underline)grandchild;
                                                                    if (und.Val == null)
                                                                    {
                                                                        bUnderline = true;
                                                                    }
                                                                    else
                                                                    {
                                                                        if (und.Val.Value != UnderlineValues.None) bUnderline = true;
                                                                        else bUnderline = false;
                                                                    }
                                                                }
                                                                else if (grandchild is FontScheme)
                                                                {
                                                                    vFontScheme = ((FontScheme)grandchild).Val;
                                                                }
                                                            }
                                                        }
                                                        #endregion

                                                        if (vFontScheme == FontSchemeValues.Major) sFontName = SimpleTheme.MajorLatinFont;
                                                        else if (vFontScheme == FontSchemeValues.Minor) sFontName = SimpleTheme.MinorLatinFont;

                                                        if (bBold) drawstyle |= System.Drawing.FontStyle.Bold;
                                                        if (bItalic) drawstyle |= System.Drawing.FontStyle.Italic;
                                                        if (bStrike) drawstyle |= System.Drawing.FontStyle.Strikeout;
                                                        if (bUnderline) drawstyle |= System.Drawing.FontStyle.Underline;

                                                        // the last element has the trailing spaces ignored. Hence the Reverse() above.
                                                        if (i == 0) sText = sText.TrimEnd();

                                                        szf = SLTool.MeasureText(bm, g, sText, SLTool.GetUsableNormalFont(sFontName, fFontSize, drawstyle, gbThrowExceptionsIfAny));
                                                        if (szf.Height > fHeight) fHeight = szf.Height;
                                                        fWidth += szf.Width;
                                                    }

                                                    ++i;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (c.DataType == CellValues.Number)
                                    {
                                        #region Numbers
                                        if (sDotNetFormatCode.Length > 0)
                                        {
                                            if (c.CellText != null)
                                            {
                                                if (!double.TryParse(c.CellText, out fCellValue))
                                                {
                                                    fCellValue = 0;
                                                }

                                                if (sFormatCode.Equals("@"))
                                                {
                                                    sText = fCellValue.ToString(CultureInfo.InvariantCulture);
                                                }
                                                else
                                                {
                                                    sText = SLTool.ToSampleDisplayFormat(fCellValue, sDotNetFormatCode, gbThrowExceptionsIfAny);
                                                }
                                            }
                                            else
                                            {
                                                if (sFormatCode.Equals("@"))
                                                {
                                                    sText = c.NumericValue.ToString(CultureInfo.InvariantCulture);
                                                }
                                                else
                                                {
                                                    sText = SLTool.ToSampleDisplayFormat(c.NumericValue, sDotNetFormatCode, gbThrowExceptionsIfAny);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (c.CellText != null)
                                            {
                                                if (!double.TryParse(c.CellText, out fCellValue))
                                                {
                                                    fCellValue = 0;
                                                }

                                                sText = SLTool.ToSampleDisplayFormat(fCellValue, "G10", gbThrowExceptionsIfAny);
                                            }
                                            else
                                            {
                                                sText = SLTool.ToSampleDisplayFormat(c.NumericValue, "G10", gbThrowExceptionsIfAny);
                                            }
                                        }
                                        #endregion
                                    }
                                    else if (c.DataType == CellValues.Boolean)
                                    {
                                        if (c.NumericValue > 0.5) sText = "TRUE";
                                        else sText = "FALSE";
                                    }
                                    else
                                    {
                                        if (c.CellText != null) sText = c.CellText;
                                        else sText = string.Empty;
                                    }

                                    sAutoFitCacheKey = string.Format("{0}{1}{2}", sText, SLConstants.AutoFitCacheSeparator, iStyleIndex.ToString(CultureInfo.InvariantCulture));
                                    if (dictAutoFitTextCache.ContainsKey(sAutoFitCacheKey))
                                    {
                                        szf = dictAutoFitTextCache[sAutoFitCacheKey];
                                        fHeight = szf.Height;
                                        fWidth = szf.Width;
                                    }
                                    else
                                    {
                                        szf = SLTool.MeasureText(bm, g, sText, ftUsable);
                                        fHeight = szf.Height;
                                        fWidth = szf.Width;
                                        dictAutoFitTextCache[sAutoFitCacheKey] = new System.Drawing.SizeF(fWidth, fHeight);
                                    }
                                }

                                if (!bSkipAdjustment)
                                {
                                    // Through empirical experimental data, it appears that there's still a bit of padding
                                    // at the end of the column when autofitting column widths. I don't know how to
                                    // calculate this padding. So I guess. I experimented with the widths of obvious
                                    // characters such as a space, an exclamation mark, a period.

                                    // Then I remember there's the documentation on the Open XML class property
                                    // Column.Width, which says there's an extra 5 pixels, 2 pixels on the left/right
                                    // and a pixel for the gridlines.

                                    // Note that this padding appears to change depending on the font/typeface and 
                                    // font size used. (Haha... where have I seen this before...) So 5 pixels doesn't
                                    // seem to work exactly. Or maybe it's wrong because the method of measuring isn't
                                    // what Excel actually uses to measure the text.

                                    // Since we're autofitting, it seems fitting (haha) that the column width is slightly
                                    // larger to accomodate the text. So it's best to err on the larger side.
                                    // Thus we add 7 instead of the "recommended" or "documented" 5 pixels, 1 extra pixel
                                    // on the left and right.
                                    fWidth += 7;
                                    // I could also have used 8, but it might have been too much of an extra padding.
                                    // The number 8 is a lucky number in Chinese culture. Goodness knows I need some
                                    // luck figuring out what Excel is doing...

                                    if (iTextRotation != 0)
                                    {
                                        szf = SLTool.CalculateOuterBoundsOfRotatedRectangle(fWidth, fHeight, iTextRotation);
                                        fHeight = szf.Height;
                                        fWidth = szf.Width;
                                    }

                                    // meaning the shared string portion was accessed (otherwise it'd be empty string)
                                    if (sAutoFitSharedStringCacheKey.Length > 0)
                                    {
                                        dictAutoFitSharedStringCache[sAutoFitSharedStringCacheKey] = new System.Drawing.SizeF(fWidth, fHeight);
                                    }
                                }

                                if (IsRow)
                                {
                                    if (fHeight > MaxPixelLength) fHeight = MaxPixelLength;

                                    if (pixellength[iPointIndex] < fHeight)
                                    {
                                        pixellength[iPointIndex] = Convert.ToInt32(Math.Ceiling(fHeight));
                                    }
                                }
                                else
                                {
                                    if (hsFilter.Contains(new SLCellPoint(rowkey, colkey))) fWidth += SLConstants.AutoFilterPixelWidth;
                                    if (fWidth > MaxPixelLength) fWidth = MaxPixelLength;

                                    if (pixellength[iPointIndex] < fWidth)
                                    {
                                        pixellength[iPointIndex] = Convert.ToInt32(Math.Ceiling(fWidth));
                                    }
                                }
                            }
                        }
                    }

                    // end of Graphics
                }
            }

            return pixellength;
        }
    }
}
