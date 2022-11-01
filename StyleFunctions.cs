using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// Get a list of existing styles. WARNING: This is only a snapshot. Any changes made to the returned result are not used.
        /// </summary>
        /// <returns>A list of existing SLStyle objects.</returns>
        public List<SLStyle> GetStyles()
        {
            List<SLStyle> result = new List<SLStyle>();
            SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            for (int i = 0; i < listStyle.Count; ++i)
            {
                style.FromHash(listStyle[i]);
                result.Add(style.Clone());
            }

            return result;
        }

        /// <summary>
        /// Get the cell's style. The default style is returned if cell doesn't have an existing style, or if the cell reference is invalid.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>The cell's style.</returns>
        public SLStyle GetCellStyle(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                style.FromHash(listStyle[0]);
                return style;
            }

            return GetCellStyle(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell's style. The default style is returned if cell doesn't have an existing style, or if the row or column indices are invalid.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The cell's style.</returns>
        public SLStyle GetCellStyle(int RowIndex, int ColumnIndex)
        {
            bool bFound = false;
            SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    bFound = true;
                    style.FromHash(listStyle[(int)c.StyleIndex]);
                }
            }

            if (!bFound)
            {
                style.FromHash(listStyle[0]);
            }

            return style;
        }

        /// <summary>
        /// Set the cell's style.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Style">The style to set.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetCellStyle(string CellReference, SLStyle Style)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return SetCellStyle(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex, Style);
        }

        /// <summary>
        /// Set the cell's style.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Style">The style to set.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetCellStyle(int RowIndex, int ColumnIndex, SLStyle Style)
        {
            return SetCellStyle(RowIndex, ColumnIndex, RowIndex, ColumnIndex, Style);
        }

        /// <summary>
        /// Set the style of a range of cells.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="Style">The style to set.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetCellStyle(string StartCellReference, string EndCellReference, SLStyle Style)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            return SetCellStyle(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Style);
        }

        /// <summary>
        /// Set the style of a range of cells.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="Style">The style to set.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetCellStyle(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLStyle Style)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
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
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit
                && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit
                && iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit
                && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                result = true;
                int iStyleIndex = this.SaveToStylesheet(Style.ToHash());

                // original style index, new style index
                Dictionary<uint, uint> stylecache = new Dictionary<uint, uint>();

                SLCell c;
                uint iCacheStyleIndex;
                SLStyle cellstyle = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                SLRowProperties rp;
                SLColumnProperties cp;
                int i, j;

                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        if (slws.CellWarehouse.Exists(i, j))
                        {
                            c = slws.CellWarehouse.Cells[i][j];
                            iCacheStyleIndex = c.StyleIndex;
                            if (stylecache.ContainsKey(iCacheStyleIndex))
                            {
                                c.StyleIndex = stylecache[iCacheStyleIndex];
                            }
                            else
                            {
                                cellstyle.FromHash(listStyle[(int)iCacheStyleIndex]);
                                cellstyle.MergeStyle(Style);
                                c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                stylecache[iCacheStyleIndex] = c.StyleIndex;
                            }
                            slws.CellWarehouse.SetValue(i, j, c);
                        }
                        else
                        {
                            if (slws.RowProperties.ContainsKey(i))
                            {
                                rp = slws.RowProperties[i];
                                iCacheStyleIndex = rp.StyleIndex;

                                c = new SLCell();
                                c.CellText = string.Empty;
                                if (stylecache.ContainsKey(iCacheStyleIndex))
                                {
                                    c.StyleIndex = stylecache[iCacheStyleIndex];
                                }
                                else
                                {
                                    cellstyle.FromHash(listStyle[(int)iCacheStyleIndex]);
                                    cellstyle.MergeStyle(Style);
                                    c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                    stylecache[iCacheStyleIndex] = c.StyleIndex;
                                }
                                slws.CellWarehouse.SetValue(i, j, c);
                            }
                            else if (slws.ColumnProperties.ContainsKey(j))
                            {
                                cp = slws.ColumnProperties[j];
                                iCacheStyleIndex = cp.StyleIndex;

                                c = new SLCell();
                                c.CellText = string.Empty;
                                if (stylecache.ContainsKey(iCacheStyleIndex))
                                {
                                    c.StyleIndex = stylecache[iCacheStyleIndex];
                                }
                                else
                                {
                                    cellstyle.FromHash(listStyle[(int)iCacheStyleIndex]);
                                    cellstyle.MergeStyle(Style);
                                    c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                    stylecache[iCacheStyleIndex] = c.StyleIndex;
                                }
                                slws.CellWarehouse.SetValue(i, j, c);
                            }
                            else
                            {
                                // it's a completely empty cell, so if it's the default style (meaning 0),
                                // there's really no point in creating a new SLCell.
                                if (iStyleIndex > 0)
                                {
                                    c = new SLCell();
                                    c.CellText = string.Empty;
                                    c.StyleIndex = (uint)iStyleIndex;
                                    slws.CellWarehouse.SetValue(i, j, c);
                                }
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Remove the style from a cell.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        public void RemoveCellStyle(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                RemoveCellStyle(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex);
            }
        }

        /// <summary>
        /// Remove the style from a range of cells.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        public void RemoveCellStyle(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                && SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                RemoveCellStyle(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
            }
        }

        /// <summary>
        /// Remove the style from a cell.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        public void RemoveCellStyle(int RowIndex, int ColumnIndex)
        {
            RemoveCellStyle(RowIndex, ColumnIndex, RowIndex, ColumnIndex);
        }

        /// <summary>
        /// Remove the style from a range of cells.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        public void RemoveCellStyle(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
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

            SLCell c;
            int i, j;

            for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
            {
                for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                {
                    if (slws.CellWarehouse.Exists(i, j))
                    {
                        c = slws.CellWarehouse.Cells[i][j];
                        c.StyleIndex = 0;
                        slws.CellWarehouse.SetValue(i, j, c);
                    }
                }
            }
        }

        /// <summary>
        /// Apply a named cell style to a cell. Existing styles are kept, unless the chosen named cell style overrides those styles.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="NamedCellStyle">The named cell style to be applied.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool ApplyNamedCellStyle(string CellReference, SLNamedCellStyleValues NamedCellStyle)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return ApplyNamedCellStyle(iRowIndex, iColumnIndex, NamedCellStyle);
        }

        /// <summary>
        /// Apply a named cell style to a cell. Existing styles are kept, unless the chosen named cell style overrides those styles.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="NamedCellStyle">The named cell style to be applied.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool ApplyNamedCellStyle(int RowIndex, int ColumnIndex, SLNamedCellStyleValues NamedCellStyle)
        {
            return ApplyNamedCellStyle(RowIndex, ColumnIndex, RowIndex, ColumnIndex, NamedCellStyle);
        }

        /// <summary>
        /// Apply a named cell style to a range of cells. Existing styles are kept, unless the chosen named cell style overrides those styles.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="NamedCellStyle">The named cell style to be applied.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool ApplyNamedCellStyle(string StartCellReference, string EndCellReference, SLNamedCellStyleValues NamedCellStyle)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            return ApplyNamedCellStyle(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, NamedCellStyle);
        }

        /// <summary>
        /// Apply a named cell style to a range of cells. Existing styles are kept, unless the chosen named cell style overrides those styles.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="StartColumnIndex">The column index of the starting column.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <param name="EndColumnIndex">The column index of the ending column.</param>
        /// <param name="NamedCellStyle">The named cell style to be applied.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool ApplyNamedCellStyle(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLNamedCellStyleValues NamedCellStyle)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            bool result = false;
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

            SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            if (SLTool.CheckRowColumnIndexLimit(iStartRowIndex, iStartColumnIndex) && SLTool.CheckRowColumnIndexLimit(iEndRowIndex, iEndColumnIndex))
            {
                result = true;
                int i = 0, j = 0;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        style = this.GetCellStyle(i, j);
                        style.ApplyNamedCellStyle(NamedCellStyle);
                        this.SetCellStyle(i, j, style);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Apply a named cell style to a row. Existing styles are kept, unless the chosen named cell style overrides those styles.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="NamedCellStyle">The named cell style to be applied.</param>
        /// <returns>True if the row index is valid. False otherwise.</returns>
        public bool ApplyNamedCellStyleToRow(int RowIndex, SLNamedCellStyleValues NamedCellStyle)
        {
            return ApplyNamedCellStyleToRow(RowIndex, RowIndex, NamedCellStyle);
        }

        /// <summary>
        /// Apply a named cell style to a range of rows. Existing styles are kept, unless the chosen named cell style overrides those styles.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <param name="NamedCellStyle">The named cell style to be applied.</param>
        /// <returns>True if the row indices are valid. False otherwise.</returns>
        public bool ApplyNamedCellStyleToRow(int StartRowIndex, int EndRowIndex, SLNamedCellStyleValues NamedCellStyle)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            bool result = false;
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

            SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                result = true;
                int i = 0;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    style = this.GetRowStyle(i);
                    style.ApplyNamedCellStyle(NamedCellStyle);
                    this.SetRowStyle(i, style);
                }
            }

            return result;
        }

        /// <summary>
        /// Apply a named cell style to a column. Existing styles are kept, unless the chosen named cell style overrides those styles.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="NamedCellStyle">The named cell style to be applied.</param>
        /// <returns>True if the column index is valid. False otherwise.</returns>
        public bool ApplyNamedCellStyleToColumn(int ColumnIndex, SLNamedCellStyleValues NamedCellStyle)
        {
            return ApplyNamedCellStyleToColumn(ColumnIndex, ColumnIndex, NamedCellStyle);
        }

        /// <summary>
        /// Apply a named cell style to a range of columns. Existing styles are kept, unless the chosen named cell style overrides those styles.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the starting column.</param>
        /// <param name="EndColumnIndex">The column index of the ending column.</param>
        /// <param name="NamedCellStyle">The named cell style to be applied.</param>
        /// <returns>True if the column indices are valid. False otherwise.</returns>
        public bool ApplyNamedCellStyleToColumn(int StartColumnIndex, int EndColumnIndex, SLNamedCellStyleValues NamedCellStyle)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            bool result = false;

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

            SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            if (iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                result = true;
                int i = 0;
                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    style = this.GetColumnStyle(i);
                    style.ApplyNamedCellStyle(NamedCellStyle);
                    this.SetColumnStyle(i, style);
                }
            }

            return result;
        }

        /// <summary>
        /// Get the style of the row. If the row doesn't have an existing style, the default style is returned.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>The row style.</returns>
        public SLStyle GetRowStyle(int RowIndex)
        {
            bool bFound = false;
            SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            if (slws.RowProperties.ContainsKey(RowIndex))
            {
                SLRowProperties rp = slws.RowProperties[RowIndex];
                bFound = true;
                style.FromHash(listStyle[(int)rp.StyleIndex]);
            }

            if (!bFound)
            {
                style.FromHash(listStyle[0]);
            }

            return style;
        }

        /// <summary>
        /// Set the row style.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="RowStyle">The style for the row.</param>
        /// <returns>True if the row index is valid. False otherwise.</returns>
        public bool SetRowStyle(int RowIndex, SLStyle RowStyle)
        {
            return SetRowStyle(RowIndex, RowIndex, RowStyle);
        }

        /// <summary>
        /// Set the row style for a range of rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        /// <param name="RowStyle">The style for the rows.</param>
        /// <returns>True if the row indices are valid. False otherwise.</returns>
        public bool SetRowStyle(int StartRowIndex, int EndRowIndex, SLStyle RowStyle)
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
                int iStyleIndex = this.SaveToStylesheet(RowStyle.ToHash());
                SLRowProperties rp;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        rp.StyleIndex = (uint)iStyleIndex;
                        slws.RowProperties[i] = rp.Clone();
                    }
                    else
                    {
                        rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                        rp.StyleIndex = (uint)iStyleIndex;
                        slws.RowProperties[i] = rp;
                    }
                }

                // original style index, new style index
                Dictionary<uint, uint> stylecache = new Dictionary<uint, uint>();

                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                List<int> listColumnKeys;
                SLCell c;
                uint iCacheStyleIndex;
                SLStyle cellstyle = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                foreach (int rowkey in listRowKeys)
                {
                    if (iStartRowIndex <= rowkey && rowkey <= iEndRowIndex)
                    {
                        listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                        foreach (int colkey in listColumnKeys)
                        {
                            c = slws.CellWarehouse.Cells[rowkey][colkey];
                            iCacheStyleIndex = c.StyleIndex;
                            if (stylecache.ContainsKey(iCacheStyleIndex))
                            {
                                c.StyleIndex = stylecache[iCacheStyleIndex];
                            }
                            else
                            {
                                cellstyle.FromHash(listStyle[(int)iCacheStyleIndex]);
                                cellstyle.MergeStyle(RowStyle);
                                c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                stylecache[iCacheStyleIndex] = c.StyleIndex;
                            }
                            slws.CellWarehouse.SetValue(rowkey, colkey, c);
                        }
                    }
                }

                List<int> colindexkeys = slws.ColumnProperties.Keys.ToList<int>();
                SLColumnProperties cp;
                foreach (int colindex in colindexkeys)
                {
                    cp = slws.ColumnProperties[colindex];
                    iCacheStyleIndex = cp.StyleIndex;
                    for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                    {
                        if (!slws.CellWarehouse.Exists(i, colindex))
                        {
                            c = new SLCell();
                            c.CellText = string.Empty;
                            if (stylecache.ContainsKey(iCacheStyleIndex))
                            {
                                c.StyleIndex = stylecache[iCacheStyleIndex];
                            }
                            else
                            {
                                cellstyle.FromHash(listStyle[(int)iCacheStyleIndex]);
                                cellstyle.MergeStyle(RowStyle);
                                c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                stylecache[iCacheStyleIndex] = c.StyleIndex;
                            }
                            slws.CellWarehouse.SetValue(i, colindex, c);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        public void DrawBorderGrid(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, b, true);
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void DrawBorderGrid(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.Color = BorderColor;
            b.BottomBorder.Color = BorderColor;
            b.LeftBorder.Color = BorderColor;
            b.RightBorder.Color = BorderColor;

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, b, true);
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        public void DrawBorderGrid(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor);
            b.BottomBorder.SetBorderThemeColor(BorderColor);
            b.LeftBorder.SetBorderThemeColor(BorderColor);
            b.RightBorder.SetBorderThemeColor(BorderColor);

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, b, true);
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void DrawBorderGrid(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor, Tint);
            b.BottomBorder.SetBorderThemeColor(BorderColor, Tint);
            b.LeftBorder.SetBorderThemeColor(BorderColor, Tint);
            b.RightBorder.SetBorderThemeColor(BorderColor, Tint);

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, b, true);
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="Border">The SLBorder object with border style properties.</param>
        public void DrawBorderGrid(string StartCellReference, string EndCellReference, SLBorder Border)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Border, true);
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        public void DrawBorderGrid(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, b, true);
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void DrawBorderGrid(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.Color = BorderColor;
            b.BottomBorder.Color = BorderColor;
            b.LeftBorder.Color = BorderColor;
            b.RightBorder.Color = BorderColor;

            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, b, true);
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        public void DrawBorderGrid(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor);
            b.BottomBorder.SetBorderThemeColor(BorderColor);
            b.LeftBorder.SetBorderThemeColor(BorderColor);
            b.RightBorder.SetBorderThemeColor(BorderColor);

            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, b, true);
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void DrawBorderGrid(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor, Tint);
            b.BottomBorder.SetBorderThemeColor(BorderColor, Tint);
            b.LeftBorder.SetBorderThemeColor(BorderColor, Tint);
            b.RightBorder.SetBorderThemeColor(BorderColor, Tint);

            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, b, true);
        }

        /// <summary>
        /// Draw border grid.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="Border">The SLBorder object with border style properties.</param>
        public void DrawBorderGrid(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLBorder Border)
        {
            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, Border, true);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        public void DrawBorder(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, b, false);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void DrawBorder(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.Color = BorderColor;
            b.BottomBorder.Color = BorderColor;
            b.LeftBorder.Color = BorderColor;
            b.RightBorder.Color = BorderColor;

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, b, false);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        public void DrawBorder(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor);
            b.BottomBorder.SetBorderThemeColor(BorderColor);
            b.LeftBorder.SetBorderThemeColor(BorderColor);
            b.RightBorder.SetBorderThemeColor(BorderColor);

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, b, false);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void DrawBorder(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor, Tint);
            b.BottomBorder.SetBorderThemeColor(BorderColor, Tint);
            b.LeftBorder.SetBorderThemeColor(BorderColor, Tint);
            b.RightBorder.SetBorderThemeColor(BorderColor, Tint);

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, b, false);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="Border">The SLBorder object with border style properties.</param>
        public void DrawBorder(string StartCellReference, string EndCellReference, SLBorder Border)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return;
            }

            DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Border, false);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        public void DrawBorder(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, b, false);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void DrawBorder(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.Color = BorderColor;
            b.BottomBorder.Color = BorderColor;
            b.LeftBorder.Color = BorderColor;
            b.RightBorder.Color = BorderColor;

            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, b, false);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        public void DrawBorder(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor);
            b.BottomBorder.SetBorderThemeColor(BorderColor);
            b.LeftBorder.SetBorderThemeColor(BorderColor);
            b.RightBorder.SetBorderThemeColor(BorderColor);

            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, b, false);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void DrawBorder(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor, Tint);
            b.BottomBorder.SetBorderThemeColor(BorderColor, Tint);
            b.LeftBorder.SetBorderThemeColor(BorderColor, Tint);
            b.RightBorder.SetBorderThemeColor(BorderColor, Tint);

            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, b, false);
        }

        /// <summary>
        /// Draw border.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="Border">The SLBorder object with border style properties.</param>
        public void DrawBorder(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLBorder Border)
        {
            DrawBorderFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, Border, false);
        }

        private void DrawBorderFinal(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLBorder Border, bool DrawBorderGrid)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
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

            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit
                && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit
                && iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit
                && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                SLStyle style;
                int i, j;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        style = this.GetCellStyle(i, j);

                        if (DrawBorderGrid)
                        {
                            style.borderReal.bpTopBorder = Border.bpTopBorder.Clone();
                            style.borderReal.bpBottomBorder = Border.bpBottomBorder.Clone();
                            style.borderReal.bpLeftBorder = Border.bpLeftBorder.Clone();
                            style.borderReal.bpRightBorder = Border.bpRightBorder.Clone();
                        }
                        else
                        {
                            if (i == iStartRowIndex)
                            {
                                style.borderReal.bpTopBorder = Border.bpTopBorder.Clone();
                            }
                            if (i == iEndRowIndex)
                            {
                                style.borderReal.bpBottomBorder = Border.bpBottomBorder.Clone();
                            }
                            if (j == iStartColumnIndex)
                            {
                                style.borderReal.bpLeftBorder = Border.bpLeftBorder.Clone();
                            }
                            if (j == iEndColumnIndex)
                            {
                                style.borderReal.bpRightBorder = Border.bpRightBorder.Clone();
                            }
                        }

                        this.SetCellStyle(i, j, style);
                    }
                }

                int iOuterTopRowIndex = iStartRowIndex - 1;
                int iOuterBottomRowIndex = iEndRowIndex + 1;
                int iOuterLeftColumnIndex = iStartColumnIndex - 1;
                int iOuterRightColumnIndex = iEndColumnIndex + 1;

                // set the bottom of the outer top row using the style of the top border
                // because the bottom of the outer top must be the same as the top of
                // the inner box.
                if (iOuterTopRowIndex >= 1)
                {
                    for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        style = this.GetCellStyle(iOuterTopRowIndex, j);
                        style.borderReal.bpBottomBorder = Border.bpTopBorder.Clone();
                        this.SetCellStyle(iOuterTopRowIndex, j, style);
                    }
                }

                // similarly for outer bottom
                if (iOuterBottomRowIndex <= SLConstants.RowLimit)
                {
                    for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        style = this.GetCellStyle(iOuterBottomRowIndex, j);
                        style.borderReal.bpTopBorder = Border.bpBottomBorder.Clone();
                        this.SetCellStyle(iOuterBottomRowIndex, j, style);
                    }
                }

                // similarly for outer left
                if (iOuterLeftColumnIndex >= 1)
                {
                    for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                    {
                        style = this.GetCellStyle(i, iOuterLeftColumnIndex);
                        style.borderReal.bpRightBorder = Border.bpLeftBorder.Clone();
                        this.SetCellStyle(i, iOuterLeftColumnIndex, style);
                    }
                }

                // similarly for outer right
                if (iOuterRightColumnIndex <= SLConstants.ColumnLimit)
                {
                    for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                    {
                        style = this.GetCellStyle(i, iOuterRightColumnIndex);
                        style.borderReal.bpLeftBorder = Border.bpRightBorder.Clone();
                        this.SetCellStyle(i, iOuterRightColumnIndex, style);
                    }
                }
            }
        }

        /// <summary>
        /// Remove any existing row style.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        public void RemoveRowStyle(int RowIndex)
        {
            RemoveRowStyle(RowIndex, RowIndex);
        }

        /// <summary>
        /// Remove any existing row style for a range of rows.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the starting row.</param>
        /// <param name="EndRowIndex">The row index of the ending row.</param>
        public void RemoveRowStyle(int StartRowIndex, int EndRowIndex)
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

            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                int i = 0;
                SLRowProperties rp;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        rp.StyleIndex = 0;
                        slws.RowProperties[i] = rp.Clone();
                    }
                }
            }
        }

        /// <summary>
        /// Get the style of the column. If the column doesn't have an existing style, the default style is returned.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The column style.</returns>
        public SLStyle GetColumnStyle(int ColumnIndex)
        {
            bool bFound = false;
            SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            if (slws.ColumnProperties.ContainsKey(ColumnIndex))
            {
                SLColumnProperties cp = slws.ColumnProperties[ColumnIndex];
                bFound = true;
                style.FromHash(listStyle[(int)cp.StyleIndex]);
            }

            if (!bFound)
            {
                style.FromHash(listStyle[0]);
            }

            return style;
        }

        /// <summary>
        /// Set the column style.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="ColumnStyle">The style for the column.</param>
        /// <returns>True if the column index is valid. False otherwise.</returns>
        public bool SetColumnStyle(int ColumnIndex, SLStyle ColumnStyle)
        {
            return SetColumnStyle(ColumnIndex, ColumnIndex, ColumnStyle);
        }

        /// <summary>
        /// Set the column style for a range of rows.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the starting column.</param>
        /// <param name="EndColumnIndex">The column index of the ending column.</param>
        /// <param name="ColumnStyle">The style for the columns.</param>
        /// <returns>True if the column indices are valid. False otherwise.</returns>
        public bool SetColumnStyle(int StartColumnIndex, int EndColumnIndex, SLStyle ColumnStyle)
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
                int iStyleIndex = this.SaveToStylesheet(ColumnStyle.ToHash());
                SLColumnProperties cp;
                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    if (slws.ColumnProperties.ContainsKey(i))
                    {
                        cp = slws.ColumnProperties[i];
                        cp.StyleIndex = (uint)iStyleIndex;
                        slws.ColumnProperties[i] = cp.Clone();
                    }
                    else
                    {
                        cp = new SLColumnProperties(SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize);
                        cp.StyleIndex = (uint)iStyleIndex;
                        slws.ColumnProperties[i] = cp;
                    }
                }

                // original style index, new style index
                Dictionary<uint, uint> stylecache = new Dictionary<uint, uint>();

                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                List<int> listColumnKeys;
                SLCell c;
                uint iCacheStyleIndex;
                SLStyle cellstyle = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        if (iStartColumnIndex <= colkey && colkey <= iEndColumnIndex)
                        {
                            c = slws.CellWarehouse.Cells[rowkey][colkey];
                            iCacheStyleIndex = c.StyleIndex;
                            if (stylecache.ContainsKey(iCacheStyleIndex))
                            {
                                c.StyleIndex = stylecache[iCacheStyleIndex];
                            }
                            else
                            {
                                cellstyle.FromHash(listStyle[(int)iCacheStyleIndex]);
                                cellstyle.MergeStyle(ColumnStyle);
                                c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                stylecache[iCacheStyleIndex] = c.StyleIndex;
                            }
                            slws.CellWarehouse.SetValue(rowkey, colkey, c);
                        }
                    }
                }

                List<int> rowindexkeys = slws.RowProperties.Keys.ToList<int>();
                SLRowProperties rp;
                SLCellPoint intersectionpt;
                foreach (int rowindex in rowindexkeys)
                {
                    rp = slws.RowProperties[rowindex];
                    iCacheStyleIndex = rp.StyleIndex;
                    for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                    {
                        intersectionpt = new SLCellPoint(rowindex, i);
                        if (!slws.CellWarehouse.Exists(rowindex, i))
                        {
                            c = new SLCell();
                            c.CellText = string.Empty;
                            if (stylecache.ContainsKey(iCacheStyleIndex))
                            {
                                c.StyleIndex = stylecache[iCacheStyleIndex];
                            }
                            else
                            {
                                cellstyle.FromHash(listStyle[(int)iCacheStyleIndex]);
                                cellstyle.MergeStyle(ColumnStyle);
                                c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                stylecache[iCacheStyleIndex] = c.StyleIndex;
                            }
                            slws.CellWarehouse.SetValue(rowindex, i, c);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Remove any existing column style.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        public void RemoveColumnStyle(int ColumnIndex)
        {
            RemoveColumnStyle(ColumnIndex, ColumnIndex);
        }

        /// <summary>
        /// Remove any existing column style for a range of columns.
        /// </summary>
        /// <param name="StartColumnIndex">The column index of the starting column.</param>
        /// <param name="EndColumnIndex">The column index of the ending column.</param>
        public void RemoveColumnStyle(int StartColumnIndex, int EndColumnIndex)
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

            if (iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                int i = 0;
                SLColumnProperties cp;
                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    if (slws.ColumnProperties.ContainsKey(i))
                    {
                        cp = slws.ColumnProperties[i];
                        cp.StyleIndex = 0;
                        slws.ColumnProperties[i] = cp.Clone();
                    }
                }
            }
        }

        /// <summary>
        /// Copy the style of one cell to another cell.
        /// </summary>
        /// <param name="FromCellReference">The cell reference of the cell whose style is copied from.</param>
        /// <param name="ToCellReference">The cell reference of the cell whose style is copied to.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellStyle(string FromCellReference, string ToCellReference)
        {
            int iFromRowIndex = -1;
            int iFromColumnIndex = -1;
            int iToRowIndex = -1;
            int iToColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(FromCellReference, out iFromRowIndex, out iFromColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(ToCellReference, out iToRowIndex, out iToColumnIndex))
            {
                return false;
            }

            return CopyCellStyle(iFromRowIndex, iFromColumnIndex, iToRowIndex, iToColumnIndex, iToRowIndex, iToColumnIndex);
        }

        /// <summary>
        /// Copy the style of one cell to a range of cells.
        /// </summary>
        /// <param name="FromCellReference">The cell reference of the cell whose style is copied from.</param>
        /// <param name="ToStartCellReference">The start cell reference of the cell range. This is typically the top-left cell.</param>
        /// <param name="ToEndCellReference">The end cell reference of the cell range. This is typically the bottom-right cell.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellStyle(string FromCellReference, string ToStartCellReference, string ToEndCellReference)
        {
            int iFromRowIndex = -1;
            int iFromColumnIndex = -1;
            int iToStartRowIndex = -1;
            int iToStartColumnIndex = -1;
            int iToEndRowIndex = -1;
            int iToEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(FromCellReference, out iFromRowIndex, out iFromColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(ToStartCellReference, out iToStartRowIndex, out iToStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(ToEndCellReference, out iToEndRowIndex, out iToEndColumnIndex))
            {
                return false;
            }

            return CopyCellStyle(iFromRowIndex, iFromColumnIndex, iToStartRowIndex, iToStartColumnIndex, iToEndRowIndex, iToEndColumnIndex);
        }

        /// <summary>
        /// Copy the style of one cell to another cell.
        /// </summary>
        /// <param name="FromRowIndex">The row index of the cell to be copied from.</param>
        /// <param name="FromColumnIndex">The column index of the cell to be copied from.</param>
        /// <param name="ToRowIndex">The row index of the cell to be copied to.</param>
        /// <param name="ToColumnIndex">The column index of the cell to be copied to.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellStyle(int FromRowIndex, int FromColumnIndex, int ToRowIndex, int ToColumnIndex)
        {
            return CopyCellStyle(FromRowIndex, FromColumnIndex, ToRowIndex, ToColumnIndex, ToRowIndex, ToColumnIndex);
        }

        /// <summary>
        /// Copy the style of one cell to a range of cells.
        /// </summary>
        /// <param name="FromRowIndex">The row index of the cell to be copied from.</param>
        /// <param name="FromColumnIndex">The column index of the cell to be copied from.</param>
        /// <param name="ToStartRowIndex">The row index of the starting cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="ToStartColumnIndex">The column index of the starting cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="ToEndRowIndex">The row index of the ending cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="ToEndColumnIndex">The column index of the ending cell of the cell range. This is typically the bottom-right cell.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellStyle(int FromRowIndex, int FromColumnIndex, int ToStartRowIndex, int ToStartColumnIndex, int ToEndRowIndex, int ToEndColumnIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            bool result = false;
            if (ToStartRowIndex < ToEndRowIndex)
            {
                iStartRowIndex = ToStartRowIndex;
                iEndRowIndex = ToEndRowIndex;
            }
            else
            {
                iStartRowIndex = ToEndRowIndex;
                iEndRowIndex = ToStartRowIndex;
            }

            if (ToStartColumnIndex < ToEndColumnIndex)
            {
                iStartColumnIndex = ToStartColumnIndex;
                iEndColumnIndex = ToEndColumnIndex;
            }
            else
            {
                iStartColumnIndex = ToEndColumnIndex;
                iEndColumnIndex = ToStartColumnIndex;
            }

            if (SLTool.CheckRowColumnIndexLimit(FromRowIndex, FromColumnIndex)
                && SLTool.CheckRowColumnIndexLimit(iStartRowIndex, iStartColumnIndex)
                && SLTool.CheckRowColumnIndexLimit(iEndRowIndex, iEndColumnIndex))
            {
                result = true;

                uint iStyleIndex = 0;
                if (slws.CellWarehouse.Exists(FromRowIndex, FromColumnIndex))
                {
                    iStyleIndex = slws.CellWarehouse.Cells[FromRowIndex][FromColumnIndex].StyleIndex;
                }
                else
                {
                    if (slws.RowProperties.ContainsKey(FromRowIndex))
                    {
                        iStyleIndex = slws.RowProperties[FromRowIndex].StyleIndex;
                    }

                    if (iStyleIndex == 0 && slws.ColumnProperties.ContainsKey(FromColumnIndex))
                    {
                        iStyleIndex = slws.ColumnProperties[FromColumnIndex].StyleIndex;
                    }
                }

                uint iStyleIndexNew;
                SLCell c;
                // we'll just overwrite any existing styles, instead of merging
                // like when we're copying row/column styles.
                for (int i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    for (int j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        // 29 Oct 2016: changed from && to ||
                        // Implemented De Morgan's law wrongly...
                        // Thanks to Christian Z for pointing this out!
                        if (i != FromRowIndex || j != FromColumnIndex)
                        {
                            if (slws.CellWarehouse.Exists(i, j))
                            {
                                slws.CellWarehouse.Cells[i][j].StyleIndex = iStyleIndex;
                            }
                            else
                            {
                                iStyleIndexNew = 0;
                                if (slws.RowProperties.ContainsKey(i)) iStyleIndexNew = slws.RowProperties[i].StyleIndex;
                                if (iStyleIndexNew == 0 && slws.ColumnProperties.ContainsKey(j)) iStyleIndexNew = slws.ColumnProperties[j].StyleIndex;

                                if (iStyleIndex != 0 || iStyleIndexNew != 0)
                                {
                                    c = new SLCell();
                                    c.CellText = string.Empty;
                                    c.StyleIndex = iStyleIndex;
                                    slws.CellWarehouse.SetValue(i, j, c);
                                }
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Copy the style of one row to another row.
        /// </summary>
        /// <param name="FromRowIndex">The row index of the row to be copied from.</param>
        /// <param name="ToRowIndex">The row index of the row to be copied to.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyRowStyle(int FromRowIndex, int ToRowIndex)
        {
            return CopyRowStyle(FromRowIndex, ToRowIndex, ToRowIndex);
        }

        /// <summary>
        /// Copy the style of one row to a range of rows.
        /// </summary>
        /// <param name="FromRowIndex">The row index of the row to be copied from.</param>
        /// <param name="ToStartRowIndex">The row index of the start row of the row range. This is typically the top row.</param>
        /// <param name="ToEndRowIndex">The row index of the end row of the row range. This is typically the bottom row.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyRowStyle(int FromRowIndex, int ToStartRowIndex, int ToEndRowIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1;
            bool result = false;
            if (ToStartRowIndex < ToEndRowIndex)
            {
                iStartRowIndex = ToStartRowIndex;
                iEndRowIndex = ToEndRowIndex;
            }
            else
            {
                iStartRowIndex = ToEndRowIndex;
                iEndRowIndex = ToStartRowIndex;
            }

            if (FromRowIndex >= 1 && FromRowIndex <= SLConstants.RowLimit
                && iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit
                && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit)
            {
                result = true;

                uint iStyleIndex = 0;
                if (slws.RowProperties.ContainsKey(FromRowIndex))
                {
                    iStyleIndex = slws.RowProperties[FromRowIndex].StyleIndex;
                }

                SLRowProperties rp;
                int i;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    if (i != FromRowIndex)
                    {
                        if (slws.RowProperties.ContainsKey(i))
                        {
                            slws.RowProperties[i].StyleIndex = iStyleIndex;
                        }
                        else
                        {
                            if (iStyleIndex != 0)
                            {
                                rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                                rp.StyleIndex = iStyleIndex;
                                slws.RowProperties[i] = rp;
                            }
                        }
                    }
                }

                #region copying cell styles
                SLCell c;
                uint iCacheStyleIndex;

                Dictionary<uint, uint> stylecache = new Dictionary<uint, uint>();

                SLStyle rowstyle = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                rowstyle.FromHash(listStyle[(int)iStyleIndex]);

                SLStyle cellstyle = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                List<int> listColumnKeys;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        if (iStartRowIndex <= rowkey && rowkey <= iEndRowIndex
                        && rowkey != FromRowIndex)
                        {
                            c = slws.CellWarehouse.Cells[rowkey][colkey];
                            iCacheStyleIndex = c.StyleIndex;
                            if (stylecache.ContainsKey(iCacheStyleIndex))
                            {
                                c.StyleIndex = stylecache[iCacheStyleIndex];
                            }
                            else
                            {
                                cellstyle.FromHash(listStyle[(int)c.StyleIndex]);
                                cellstyle.MergeStyle(rowstyle);
                                c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                stylecache[iCacheStyleIndex] = c.StyleIndex;
                            }
                            slws.CellWarehouse.SetValue(rowkey, colkey, c);
                        }
                    }
                }

                // this follows the algorithm in setting row/column style.
                // See appropriate function and make sure to sync with that function.
                List<int> colindexkeys = slws.ColumnProperties.Keys.ToList<int>();
                SLColumnProperties cp;
                SLCellPoint intersectionpt;
                foreach (int colindex in colindexkeys)
                {
                    cp = slws.ColumnProperties[colindex];
                    iCacheStyleIndex = cp.StyleIndex;
                    for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                    {
                        intersectionpt = new SLCellPoint(i, colindex);
                        if (!slws.CellWarehouse.Exists(i, colindex))
                        {
                            c = new SLCell();
                            c.CellText = string.Empty;
                            if (stylecache.ContainsKey(iCacheStyleIndex))
                            {
                                c.StyleIndex = stylecache[iCacheStyleIndex];
                            }
                            else
                            {
                                cellstyle.FromHash(listStyle[(int)iCacheStyleIndex]);
                                cellstyle.MergeStyle(rowstyle);
                                c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                stylecache[iCacheStyleIndex] = c.StyleIndex;
                            }
                            slws.CellWarehouse.SetValue(i, colindex, c);
                        }
                    }
                }
                #endregion
            }

            return result;
        }

        /// <summary>
        /// Copy the style of one column to another column.
        /// </summary>
        /// <param name="FromColumnIndex">The column index of the column to be copied from.</param>
        /// <param name="ToColumnIndex">The column index of the column to be copied to.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumnStyle(int FromColumnIndex, int ToColumnIndex)
        {
            return CopyColumnStyle(FromColumnIndex, ToColumnIndex, ToColumnIndex);
        }

        /// <summary>
        /// Copy the style of one column to a range of columns.
        /// </summary>
        /// <param name="FromColumnIndex">The column index of the column to be copied from.</param>
        /// <param name="ToStartColumnIndex">The column index of the start column of the column range. This is typically the left-most column.</param>
        /// <param name="ToEndColumnIndex">The column index of the end column of the column range. This is typically the right-most column.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyColumnStyle(int FromColumnIndex, int ToStartColumnIndex, int ToEndColumnIndex)
        {
            int iStartColumnIndex = 1, iEndColumnIndex = 1;
            bool result = false;

            if (ToStartColumnIndex < ToEndColumnIndex)
            {
                iStartColumnIndex = ToStartColumnIndex;
                iEndColumnIndex = ToEndColumnIndex;
            }
            else
            {
                iStartColumnIndex = ToEndColumnIndex;
                iEndColumnIndex = ToStartColumnIndex;
            }

            if (FromColumnIndex >= 1 && FromColumnIndex <= SLConstants.ColumnLimit
                && iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit
                && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit)
            {
                result = true;

                uint iStyleIndex = 0;
                if (slws.ColumnProperties.ContainsKey(FromColumnIndex))
                {
                    iStyleIndex = slws.ColumnProperties[FromColumnIndex].StyleIndex;
                }

                SLColumnProperties cp;
                int i;
                for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                {
                    if (i != FromColumnIndex)
                    {
                        if (slws.ColumnProperties.ContainsKey(i))
                        {
                            slws.ColumnProperties[i].StyleIndex = iStyleIndex;
                        }
                        else
                        {
                            if (iStyleIndex != 0)
                            {
                                cp = new SLColumnProperties(SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize);
                                cp.StyleIndex = iStyleIndex;
                                slws.ColumnProperties[i] = cp;
                            }
                        }
                    }
                }

                #region copying cell styles
                SLCell c;
                uint iCacheStyleIndex;

                Dictionary<uint, uint> stylecache = new Dictionary<uint, uint>();

                SLStyle colstyle = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                colstyle.FromHash(listStyle[(int)iStyleIndex]);

                SLStyle cellstyle = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                List<int> listColumnKeys;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        if (iStartColumnIndex <= colkey && colkey <= iEndColumnIndex
                        && colkey != FromColumnIndex)
                        {
                            c = slws.CellWarehouse.Cells[rowkey][colkey];
                            iCacheStyleIndex = c.StyleIndex;
                            if (stylecache.ContainsKey(iCacheStyleIndex))
                            {
                                c.StyleIndex = stylecache[iCacheStyleIndex];
                            }
                            else
                            {
                                cellstyle.FromHash(listStyle[(int)c.StyleIndex]);
                                cellstyle.MergeStyle(colstyle);
                                c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                stylecache[iCacheStyleIndex] = c.StyleIndex;
                            }
                            slws.CellWarehouse.SetValue(rowkey, colkey, c);
                        }
                    }
                }

                // this follows the algorithm in setting row/column style.
                // See appropriate function and make sure to sync with that function.
                List<int> rowindexkeys = slws.RowProperties.Keys.ToList<int>();
                SLRowProperties rp;
                foreach (int rowindex in rowindexkeys)
                {
                    rp = slws.RowProperties[rowindex];
                    iCacheStyleIndex = rp.StyleIndex;
                    for (i = iStartColumnIndex; i <= iEndColumnIndex; ++i)
                    {
                        if (!slws.CellWarehouse.Exists(rowindex, i))
                        {
                            c = new SLCell();
                            c.CellText = string.Empty;
                            if (stylecache.ContainsKey(iCacheStyleIndex))
                            {
                                c.StyleIndex = stylecache[iCacheStyleIndex];
                            }
                            else
                            {
                                cellstyle.FromHash(listStyle[(int)iCacheStyleIndex]);
                                cellstyle.MergeStyle(colstyle);
                                c.StyleIndex = (uint)this.SaveToStylesheet(cellstyle.ToHash());
                                stylecache[iCacheStyleIndex] = c.StyleIndex;
                            }
                            slws.CellWarehouse.SetValue(rowindex, i, c);
                        }
                    }
                }
                #endregion
            }

            return result;
        }

        internal void TranslateStyleIdsToStyles(ref SLStyle style)
        {
            int index;

            if (style.NumberFormatId != null)
            {
                index = (int)style.NumberFormatId.Value;
                style.nfFormatCode = new SLNumberingFormat();
                style.nfFormatCode.NumberFormatId = (uint)index;

                if (dictStyleNumberingFormat.ContainsKey(index))
                {
                    style.nfFormatCode.FromHash(dictStyleNumberingFormat[index]);
                }
                else if (dictBuiltInNumberingFormat.ContainsKey(index))
                {
                    style.nfFormatCode.FormatCode = dictBuiltInNumberingFormat[index];
                }
                else
                {
                    // don't know the format code, but *something* has to be written.
                    //style.nfFormatCode.FormatCode = string.Format("Built-in format code {0}", index);
                    style.nfFormatCode.FormatCode = SLConstants.NumberFormatGeneral;
                }
                style.HasNumberingFormat = true;
            }
            else
            {
                style.RemoveFormatCode();
            }

            if (style.FontId != null)
            {
                style.HasFont = true;
                index = (int)style.FontId.Value;
                if (index >= 0 && index < listStyleFont.Count)
                {
                    style.fontReal = new SLFont(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    style.fontReal.FromHash(listStyleFont[index]);
                }
                else
                {
                    style.RemoveFont();
                }
            }
            else
            {
                style.RemoveFont();
            }

            if (style.FillId != null)
            {
                style.HasFill = true;
                index = (int)style.FillId.Value;
                if (index >= 0 && index < listStyleFill.Count)
                {
                    style.fillReal = new SLFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    style.fillReal.FromHash(listStyleFill[index]);
                }
                else
                {
                    style.RemoveFill();
                }
            }
            else
            {
                style.RemoveFill();
            }

            if (style.BorderId != null)
            {
                style.HasBorder = true;
                index = (int)style.BorderId.Value;
                if (index >= 0 && index < listStyleBorder.Count)
                {
                    style.borderReal = new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    style.borderReal.FromHash(listStyleBorder[index]);
                }
                else
                {
                    style.RemoveBorder();
                }
            }
            else
            {
                style.RemoveBorder();
            }

            style.Sync();
        }

        internal void TranslateStylesToStyleIds(ref SLStyle style)
        {
            style.Sync();

            if (style.nfFormatCode.FormatCode.Length > 0)
            {
                style.NumberFormatId = (uint)this.SaveToStylesheetNumberingFormat(style.nfFormatCode.ToHash());
            }
            else
            {
                style.NumberFormatId = (uint)this.NumberFormatGeneralId;
            }

            if (style.HasFont)
            {
                style.FontId = (uint)this.SaveToStylesheetFont(style.fontReal.ToHash());
            }
            else
            {
                style.FontId = (uint)this.SaveToStylesheetFont(listStyleFont[0]);
            }

            if (style.HasFill)
            {
                style.FillId = (uint)this.SaveToStylesheetFill(style.fillReal.ToHash());
            }
            else
            {
                SLFill fill = new SLFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                fill.SetPatternType(PatternValues.None);
                style.FillId = (uint)this.SaveToStylesheetFill(fill.ToHash());
            }

            if (style.HasBorder)
            {
                style.BorderId = (uint)this.SaveToStylesheetBorder(style.borderReal.ToHash());
            }
            else
            {
                SLBorder border = new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                style.BorderId = (uint)this.SaveToStylesheetBorder(border.ToHash());
            }
        }

        internal void LoadStylesheet()
        {
            countStyle = 0;
            listStyle = new List<string>();
            dictStyleHash = new Dictionary<string, int>();

            NumberFormatGeneralId = -1;
            NumberFormatGeneralText = SLConstants.NumberFormatGeneral;
            NextNumberFormatId = SLConstants.CustomNumberFormatIdStartIndex;
            dictStyleNumberingFormat = new Dictionary<int, string>();
            dictStyleNumberingFormatHash = new Dictionary<string, int>();

            countStyleFont = 0;
            listStyleFont = new List<string>();
            dictStyleFontHash = new Dictionary<string, int>();

            countStyleFill = 0;
            listStyleFill = new List<string>();
            dictStyleFillHash = new Dictionary<string, int>();

            countStyleBorder = 0;
            listStyleBorder = new List<string>();
            dictStyleBorderHash = new Dictionary<string, int>();

            countStyleCellStyle = 0;
            listStyleCellStyle = new List<string>();
            dictStyleCellStyleHash = new Dictionary<string, int>();

            countStyleCellStyleFormat = 0;
            listStyleCellStyleFormat = new List<string>();
            dictStyleCellStyleFormatHash = new Dictionary<string, int>();

            countStyleDifferentialFormat = 0;
            listStyleDifferentialFormat = new List<string>();
            dictStyleDifferentialFormatHash = new Dictionary<string, int>();

            countStyleTableStyle = 0;
            listStyleTableStyle = new List<string>();
            dictStyleTableStyleHash = new Dictionary<string, int>();

            int i = 0;
            string sHash = string.Empty;

            if (wbp.WorkbookStylesPart != null)
            {
                WorkbookStylesPart wbsp = wbp.WorkbookStylesPart;

                this.NextNumberFormatId = SLConstants.CustomNumberFormatIdStartIndex;
                this.StylesheetColors = null;

                using (OpenXmlReader oxr = OpenXmlReader.Create(wbsp))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(NumberingFormats))
                        {
                            NumberingFormat nf;
                            using (OpenXmlReader oxrNF = OpenXmlReader.Create((NumberingFormats)oxr.LoadCurrentElement()))
                            {
                                while (oxrNF.Read())
                                {
                                    if (oxrNF.ElementType == typeof(NumberingFormat))
                                    {
                                        nf = (NumberingFormat)oxrNF.LoadCurrentElement();
                                        if (nf.NumberFormatId != null && nf.FormatCode != null)
                                        {
                                            i = (int)nf.NumberFormatId.Value;
                                            sHash = nf.FormatCode.Value;
                                            dictStyleNumberingFormat[i] = sHash;
                                            dictStyleNumberingFormatHash[sHash] = i;

                                            if (sHash.Equals(SLConstants.NumberFormatGeneral, StringComparison.OrdinalIgnoreCase))
                                            {
                                                this.NumberFormatGeneralText = sHash;
                                                this.NumberFormatGeneralId = i;
                                            }

                                            // if there's a number format greater than the next number,
                                            // obviously we want to increment.
                                            // if there's a current number equal to the next number,
                                            // we want to increment because this number exists!
                                            // Emphasis on *next* number
                                            if (i >= this.NextNumberFormatId)
                                            {
                                                this.NextNumberFormatId = this.NextNumberFormatId + 1;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (oxr.ElementType == typeof(Fonts))
                        {
                            SLFont fontSL;
                            using (OpenXmlReader oxrFont = OpenXmlReader.Create((Fonts)oxr.LoadCurrentElement()))
                            {
                                while (oxrFont.Read())
                                {
                                    if (oxrFont.ElementType == typeof(Font))
                                    {
                                        fontSL = new SLFont(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                                        fontSL.FromFont((Font)oxrFont.LoadCurrentElement());
                                        this.ForceSaveToStylesheetFont(fontSL.ToHash());
                                    }
                                }
                            }
                            countStyleFont = listStyleFont.Count;
                        }
                        else if (oxr.ElementType == typeof(Fills))
                        {
                            SLFill fillSL;
                            using (OpenXmlReader oxrFill = OpenXmlReader.Create((Fills)oxr.LoadCurrentElement()))
                            {
                                while (oxrFill.Read())
                                {
                                    if (oxrFill.ElementType == typeof(Fill))
                                    {
                                        fillSL = new SLFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                                        fillSL.FromFill((Fill)oxrFill.LoadCurrentElement());
                                        this.ForceSaveToStylesheetFill(fillSL.ToHash());
                                    }
                                }
                            }
                            countStyleFill = listStyleFill.Count;
                        }
                        else if (oxr.ElementType == typeof(Borders))
                        {
                            SLBorder borderSL;
                            using (OpenXmlReader oxrBorder = OpenXmlReader.Create((Borders)oxr.LoadCurrentElement()))
                            {
                                while (oxrBorder.Read())
                                {
                                    if (oxrBorder.ElementType == typeof(Border))
                                    {
                                        borderSL = new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                                        borderSL.FromBorder((Border)oxrBorder.LoadCurrentElement());
                                        this.ForceSaveToStylesheetBorder(borderSL.ToHash());
                                    }
                                }
                            }
                            countStyleBorder = listStyleBorder.Count;
                        }
                        else if (oxr.ElementType == typeof(CellStyleFormats))
                        {
                            SLStyle styleSL;
                            using (OpenXmlReader oxrCellStyleFormats = OpenXmlReader.Create((CellStyleFormats)oxr.LoadCurrentElement()))
                            {
                                while (oxrCellStyleFormats.Read())
                                {
                                    if (oxrCellStyleFormats.ElementType == typeof(CellFormat))
                                    {
                                        styleSL = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                                        styleSL.FromCellFormat((CellFormat)oxrCellStyleFormats.LoadCurrentElement());
                                        this.TranslateStyleIdsToStyles(ref styleSL);
                                        this.ForceSaveToStylesheetCellStylesFormat(styleSL.ToHash());
                                    }
                                }
                            }
                            countStyleCellStyleFormat = listStyleCellStyleFormat.Count;
                        }
                        else if (oxr.ElementType == typeof(CellFormats))
                        {
                            SLStyle styleSL;
                            using (OpenXmlReader oxrCellFormats = OpenXmlReader.Create((CellFormats)oxr.LoadCurrentElement()))
                            {
                                while (oxrCellFormats.Read())
                                {
                                    if (oxrCellFormats.ElementType == typeof(CellFormat))
                                    {
                                        styleSL = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                                        styleSL.FromCellFormat((CellFormat)oxrCellFormats.LoadCurrentElement());
                                        this.TranslateStyleIdsToStyles(ref styleSL);
                                        this.ForceSaveToStylesheet(styleSL.ToHash());
                                    }
                                }
                            }
                            countStyle = listStyle.Count;
                        }
                        else if (oxr.ElementType == typeof(CellStyles))
                        {
                            SLCellStyle csSL;
                            using (OpenXmlReader oxrCellStyles = OpenXmlReader.Create((CellStyles)oxr.LoadCurrentElement()))
                            {
                                while (oxrCellStyles.Read())
                                {
                                    if (oxrCellStyles.ElementType == typeof(CellStyle))
                                    {
                                        csSL = new SLCellStyle();
                                        csSL.FromCellStyle((CellStyle)oxrCellStyles.LoadCurrentElement());
                                        this.ForceSaveToStylesheetCellStyle(csSL.ToHash());
                                    }
                                }
                            }
                            countStyleCellStyle = listStyleCellStyle.Count;
                        }
                        else if (oxr.ElementType == typeof(DifferentialFormats))
                        {
                            SLDifferentialFormat dfSL;
                            using (OpenXmlReader oxrDiff = OpenXmlReader.Create((DifferentialFormats)oxr.LoadCurrentElement()))
                            {
                                while (oxrDiff.Read())
                                {
                                    if (oxrDiff.ElementType == typeof(DifferentialFormat))
                                    {
                                        dfSL = new SLDifferentialFormat();
                                        dfSL.FromDifferentialFormat((DifferentialFormat)oxrDiff.LoadCurrentElement());
                                        this.ForceSaveToStylesheetDifferentialFormat(dfSL.ToHash());
                                    }
                                }
                            }
                            countStyleDifferentialFormat = listStyleDifferentialFormat.Count;
                        }
                        else if (oxr.ElementType == typeof(TableStyles))
                        {
                            TableStyles tss = (TableStyles)oxr.LoadCurrentElement();
                            SLTableStyle tsSL;
                            i = 0;
                            using (OpenXmlReader oxrTableStyles = OpenXmlReader.Create(tss))
                            {
                                while (oxrTableStyles.Read())
                                {
                                    if (oxrTableStyles.ElementType == typeof(TableStyle))
                                    {
                                        tsSL = new SLTableStyle();
                                        tsSL.FromTableStyle((TableStyle)oxrTableStyles.LoadCurrentElement());
                                        sHash = tsSL.ToHash();
                                        listStyleTableStyle.Add(sHash);
                                        dictStyleTableStyleHash[sHash] = i;
                                        ++i;
                                    }
                                }
                            }
                            countStyleTableStyle = listStyleTableStyle.Count;

                            if (tss.DefaultTableStyle != null)
                            {
                                this.TableStylesDefaultTableStyle = tss.DefaultTableStyle.Value;
                            }
                            else
                            {
                                this.TableStylesDefaultTableStyle = string.Empty;
                            }

                            if (tss.DefaultPivotStyle != null)
                            {
                                this.TableStylesDefaultPivotStyle = tss.DefaultPivotStyle.Value;
                            }
                            else
                            {
                                this.TableStylesDefaultPivotStyle = string.Empty;
                            }
                        }
                        else if (oxr.ElementType == typeof(Colors))
                        {
                            this.StylesheetColors = (Colors)(oxr.LoadCurrentElement().CloneNode(true));
                        }
                    }
                }

                // Force a "General" number format to be saved.
                // Upper case is used by LibreOffice. Is it case insensitive?
                if (this.NumberFormatGeneralId < 0)
                {
                    if (!dictStyleNumberingFormat.ContainsKey(0))
                    {
                        this.NumberFormatGeneralId = 0;
                        this.NumberFormatGeneralText = SLConstants.NumberFormatGeneral;
                        dictStyleNumberingFormat[this.NumberFormatGeneralId] = this.NumberFormatGeneralText;
                        dictStyleNumberingFormatHash[this.NumberFormatGeneralText] = this.NumberFormatGeneralId;
                    }
                    else
                    {
                        this.NumberFormatGeneralId = this.NextNumberFormatId;
                        this.NumberFormatGeneralText = SLConstants.NumberFormatGeneral;
                        dictStyleNumberingFormat[this.NumberFormatGeneralId] = this.NumberFormatGeneralText;
                        dictStyleNumberingFormatHash[this.NumberFormatGeneralText] = this.NumberFormatGeneralId;

                        ++this.NextNumberFormatId;
                    }
                }

                if (listStyleFont.Count == 0)
                {
                    SLFont fontSL = new SLFont(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    fontSL.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                    fontSL.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.SaveToStylesheetFont(fontSL.ToHash());
                }

                if (listStyleFill.Count == 0)
                {
                    SLFill fillNone = new SLFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    fillNone.SetPatternType(PatternValues.None);
                    this.SaveToStylesheetFill(fillNone.ToHash());

                    SLFill fillGray125 = new SLFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    fillGray125.SetPatternType(PatternValues.Gray125);
                    this.SaveToStylesheetFill(fillGray125.ToHash());
                }
                else
                {
                    // make sure there's at least a "none" pattern
                    SLFill fillNone = new SLFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    fillNone.SetPatternType(PatternValues.None);
                    this.SaveToStylesheetFill(fillNone.ToHash());
                }

                // make sure there's at least an empty border
                SLBorder borderEmpty = new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                this.SaveToStylesheetBorder(borderEmpty.ToHash());

                int iCanonicalCellStyleFormatId = 0;
                SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                style.FormatCode = this.NumberFormatGeneralText;
                style.fontReal.FromHash(listStyleFont[0]);
                style.Fill.SetPatternType(PatternValues.None);
                style.Border = new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);

                // there's at least one cell style format
                iCanonicalCellStyleFormatId = this.SaveToStylesheetCellStylesFormat(style.ToHash());

                // there's at least one style
                style.CellStyleFormatId = (uint)iCanonicalCellStyleFormatId;
                this.SaveToStylesheet(style.ToHash());

                if (listStyleCellStyle.Count == 0)
                {
                    SLCellStyle csNormal = new SLCellStyle();
                    csNormal.Name = "Normal";
                    //csNormal.FormatId = 0;
                    csNormal.FormatId = (uint)iCanonicalCellStyleFormatId;
                    csNormal.BuiltinId = 0;
                    this.SaveToStylesheetCellStyle(csNormal.ToHash());
                }
            }
            else
            {
                // no numbering format by default
                this.NextNumberFormatId = SLConstants.CustomNumberFormatIdStartIndex;

                this.NumberFormatGeneralId = 0;
                this.NumberFormatGeneralText = SLConstants.NumberFormatGeneral;
                dictStyleNumberingFormat[this.NumberFormatGeneralId] = this.NumberFormatGeneralText;
                dictStyleNumberingFormatHash[this.NumberFormatGeneralText] = this.NumberFormatGeneralId;

                SLFont fontDefault = new SLFont(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                fontDefault.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);
                fontDefault.SetFontThemeColor(SLThemeColorIndexValues.Dark1Color);
                this.SaveToStylesheetFont(fontDefault.ToHash());

                SLFill fillNone = new SLFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                fillNone.SetPatternType(PatternValues.None);
                this.SaveToStylesheetFill(fillNone.ToHash());

                SLFill fillGray125 = new SLFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                fillGray125.SetPatternType(PatternValues.Gray125);
                this.SaveToStylesheetFill(fillGray125.ToHash());

                SLBorder borderEmpty = new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                this.SaveToStylesheetBorder(borderEmpty.ToHash());

                int iCanonicalCellStyleFormatId = 0;
                SLStyle style = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                style.FormatCode = this.NumberFormatGeneralText;
                style.Font = fontDefault;
                style.Fill = fillNone;
                style.Border = borderEmpty;
                iCanonicalCellStyleFormatId = this.SaveToStylesheetCellStylesFormat(style.ToHash());

                style.CellStyleFormatId = (uint)iCanonicalCellStyleFormatId;
                this.SaveToStylesheet(style.ToHash());

                SLCellStyle csNormal = new SLCellStyle();
                csNormal.Name = "Normal";
                //csNormal.FormatId = 0;
                csNormal.FormatId = (uint)iCanonicalCellStyleFormatId;
                csNormal.BuiltinId = 0;
                this.SaveToStylesheetCellStyle(csNormal.ToHash());

                this.TableStylesDefaultTableStyle = SLConstants.DefaultTableStyle;
                this.TableStylesDefaultPivotStyle = SLConstants.DefaultPivotStyle;
            }
        }

        internal void WriteStylesheet()
        {
            int i = 0;

            SLStyle styleSL = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            for (i = 0; i < listStyle.Count; ++i)
            {
                styleSL.FromHash(listStyle[i]);
                if (styleSL.nfFormatCode.FormatCode.Length > 0)
                {
                    this.SaveToStylesheetNumberingFormat(styleSL.nfFormatCode.ToHash());
                }

                if (styleSL.HasFont)
                {
                    this.SaveToStylesheetFont(styleSL.fontReal.ToHash());
                }

                if (styleSL.HasFill)
                {
                    this.SaveToStylesheetFill(styleSL.fillReal.ToHash());
                }

                if (styleSL.HasBorder)
                {
                    this.SaveToStylesheetBorder(styleSL.borderReal.ToHash());
                }
            }

            if (wbp.WorkbookStylesPart != null)
            {
                int diff = 0;
                WorkbookStylesPart wbsp = wbp.WorkbookStylesPart;

                if (dictStyleNumberingFormat.Count > 0)
                {
                    if (dictStyleNumberingFormat.Count == 1
                        && dictStyleNumberingFormatHash.ContainsKey(SLConstants.NumberFormatGeneral)
                        && dictStyleNumberingFormatHash[SLConstants.NumberFormatGeneral] == 0)
                    {
                        // it just contains our "default" number format, so don't do anything.
                    }
                    else
                    {
                        wbsp.Stylesheet.NumberingFormats = new NumberingFormats();
                        List<int> listKeys = dictStyleNumberingFormat.Keys.ToList<int>();
                        listKeys.Sort();
                        if (listKeys[0] == 0
                            && dictStyleNumberingFormat[listKeys[0]].Equals(SLConstants.NumberFormatGeneral, StringComparison.OrdinalIgnoreCase))
                        {
                            listKeys.RemoveAt(0);
                        }
                        wbsp.Stylesheet.NumberingFormats.Count = (uint)listKeys.Count;
                        SLNumberingFormat nfSL;
                        for (i = 0; i < listKeys.Count; ++i)
                        {
                            nfSL = new SLNumberingFormat();
                            nfSL.FromHash(dictStyleNumberingFormat[listKeys[i]]);
                            wbsp.Stylesheet.NumberingFormats.Append(new NumberingFormat()
                            {
                                NumberFormatId = (uint)listKeys[i],
                                FormatCode = nfSL.FormatCode

                                // no need to escape (particular about the double quotes)
                                //FormatCode = SLTool.XmlWrite(nfSL.FormatCode)
                            });
                        }
                    }
                }

                if (listStyleFont.Count > countStyleFont)
                {
                    if (wbsp.Stylesheet.Fonts == null)
                    {
                        wbsp.Stylesheet.Fonts = new Fonts();
                    }
                    wbsp.Stylesheet.Fonts.Count = (uint)listStyleFont.Count;
                    diff = listStyleFont.Count - countStyleFont;
                    for (i = 0; i < diff; ++i)
                    {
                        wbsp.Stylesheet.Fonts.Append(new Font() { InnerXml = listStyleFont[i + countStyleFont] });
                    }
                }

                if (listStyleFill.Count > countStyleFill)
                {
                    if (wbsp.Stylesheet.Fills == null)
                    {
                        wbsp.Stylesheet.Fills = new Fills();
                    }
                    wbsp.Stylesheet.Fills.Count = (uint)listStyleFill.Count;
                    diff = listStyleFill.Count - countStyleFill;
                    for (i = 0; i < diff; ++i)
                    {
                        wbsp.Stylesheet.Fills.Append(new Fill() { InnerXml = listStyleFill[i + countStyleFill] });
                    }
                }

                if (listStyleBorder.Count > countStyleBorder)
                {
                    if (wbsp.Stylesheet.Borders == null)
                    {
                        wbsp.Stylesheet.Borders = new Borders();
                    }
                    wbsp.Stylesheet.Borders.Count = (uint)listStyleBorder.Count;
                    diff = listStyleBorder.Count - countStyleBorder;
                    SLBorder borderSL = new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    for (i = 0; i < diff; ++i)
                    {
                        borderSL.FromHash(listStyleBorder[i + countStyleBorder]);
                        wbsp.Stylesheet.Borders.Append(borderSL.ToBorder());
                    }
                }

                int iCanonicalCellStyleFormatId = 0;
                SLStyle styleCanonical = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                styleCanonical.FormatCode = this.NumberFormatGeneralText;
                styleCanonical.fontReal.FromHash(listStyleFont[0]);
                styleCanonical.Fill.SetPatternType(PatternValues.None);
                styleCanonical.Border = new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                iCanonicalCellStyleFormatId = this.SaveToStylesheetCellStylesFormat(styleCanonical.ToHash());
                this.TranslateStylesToStyleIds(ref styleCanonical);

                if (listStyleCellStyleFormat.Count > countStyleCellStyleFormat)
                {
                    if (wbsp.Stylesheet.CellStyleFormats == null)
                    {
                        wbsp.Stylesheet.CellStyleFormats = new CellStyleFormats();
                    }
                    wbsp.Stylesheet.CellStyleFormats.Count = (uint)listStyleCellStyleFormat.Count;
                    diff = listStyleCellStyleFormat.Count - countStyleCellStyleFormat;
                    styleSL = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    for (i = 0; i < diff; ++i)
                    {
                        styleSL.FromHash(listStyleCellStyleFormat[i + countStyleCellStyleFormat]);
                        this.TranslateStylesToStyleIds(ref styleSL);
                        wbsp.Stylesheet.CellStyleFormats.Append(styleSL.ToCellFormat());
                    }
                }

                if (listStyle.Count > countStyle)
                {
                    if (wbsp.Stylesheet.CellFormats == null)
                    {
                        wbsp.Stylesheet.CellFormats = new CellFormats();
                    }
                    wbsp.Stylesheet.CellFormats.Count = (uint)listStyle.Count;
                    diff = listStyle.Count - countStyle;
                    styleSL = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                    for (i = 0; i < diff; ++i)
                    {
                        styleSL.FromHash(listStyle[i + countStyle]);
                        this.TranslateStylesToStyleIds(ref styleSL);
                        styleSL.CellStyleFormatId = (uint)iCanonicalCellStyleFormatId;
                        if (styleSL.NumberFormatId != styleCanonical.NumberFormatId) styleSL.ApplyNumberFormat = true;
                        if (styleSL.FontId != styleCanonical.FontId) styleSL.ApplyFont = true;
                        if (styleSL.FillId != styleCanonical.FillId) styleSL.ApplyFill = true;
                        if (styleSL.BorderId != styleCanonical.BorderId) styleSL.ApplyBorder = true;
                        if (styleSL.HasAlignment) styleSL.ApplyAlignment = true;
                        if (styleSL.HasProtection) styleSL.ApplyProtection = true;
                        wbsp.Stylesheet.CellFormats.Append(styleSL.ToCellFormat());
                    }
                }

                if (listStyleCellStyle.Count > countStyleCellStyle)
                {
                    if (wbsp.Stylesheet.CellStyles == null)
                    {
                        wbsp.Stylesheet.CellStyles = new CellStyles();
                    }
                    wbsp.Stylesheet.CellStyles.Count = (uint)listStyleCellStyle.Count;
                    diff = listStyleCellStyle.Count - countStyleCellStyle;
                    SLCellStyle csSL = new SLCellStyle();
                    for (i = 0; i < diff; ++i)
                    {
                        csSL.FromHash(listStyleCellStyle[i + countStyleCellStyle]);
                        wbsp.Stylesheet.CellStyles.Append(csSL.ToCellStyle());
                    }
                }

                if (listStyleDifferentialFormat.Count > countStyleDifferentialFormat)
                {
                    if (wbsp.Stylesheet.DifferentialFormats == null)
                    {
                        wbsp.Stylesheet.DifferentialFormats = new DifferentialFormats();
                    }
                    wbsp.Stylesheet.DifferentialFormats.Count = (uint)listStyleDifferentialFormat.Count;
                    diff = listStyleDifferentialFormat.Count - countStyleDifferentialFormat;
                    for (i = 0; i < diff; ++i)
                    {
                        wbsp.Stylesheet.DifferentialFormats.Append(new DifferentialFormat() { InnerXml = listStyleDifferentialFormat[i + countStyleDifferentialFormat] });
                    }
                }

                if (listStyleTableStyle.Count > countStyleTableStyle)
                {
                    if (wbsp.Stylesheet.TableStyles == null)
                    {
                        wbsp.Stylesheet.TableStyles = new TableStyles();
                    }
                    wbsp.Stylesheet.TableStyles.Count = (uint)listStyleTableStyle.Count;
                    if (this.TableStylesDefaultTableStyle.Length > 0)
                    {
                        wbsp.Stylesheet.TableStyles.DefaultTableStyle = this.TableStylesDefaultTableStyle;
                    }
                    if (this.TableStylesDefaultPivotStyle.Length > 0)
                    {
                        wbsp.Stylesheet.TableStyles.DefaultPivotStyle = this.TableStylesDefaultPivotStyle;
                    }
                    diff = listStyleTableStyle.Count - countStyleTableStyle;
                    SLTableStyle tsSL = new SLTableStyle();
                    for (i = 0; i < diff; ++i)
                    {
                        tsSL.FromHash(listStyleTableStyle[i + countStyleTableStyle]);
                        wbsp.Stylesheet.TableStyles.Append(tsSL.ToTableStyle());
                    }
                }

                // we're not touching Colors here, so there shouldn't be anything to update

                wbsp.Stylesheet.Save();
            }
            else
            {
                WorkbookStylesPart wbsp = wbp.AddNewPart<WorkbookStylesPart>();
                using (MemoryStream ms = new MemoryStream())
                {
                    using (StreamWriter sw = new StreamWriter(ms))
                    {
                        sw.Write("<x:styleSheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

                        if (dictStyleNumberingFormat.Count > 0)
                        {
                            if (dictStyleNumberingFormat.Count == 1 
                                && dictStyleNumberingFormatHash.ContainsKey(SLConstants.NumberFormatGeneral)
                                && dictStyleNumberingFormatHash[SLConstants.NumberFormatGeneral] == 0)
                            {
                                // it just contains our "default" number format, so don't do anything.
                            }
                            else
                            {
                                List<int> listKeys = dictStyleNumberingFormat.Keys.ToList<int>();
                                listKeys.Sort();
                                if (listKeys[0] == 0
                                    && dictStyleNumberingFormat[listKeys[0]].Equals(SLConstants.NumberFormatGeneral, StringComparison.OrdinalIgnoreCase))
                                {
                                    listKeys.RemoveAt(0);
                                }
                                sw.Write("<x:numFmts count=\"{0}\">", listKeys.Count);
                                for (i = 0; i < listKeys.Count; ++i)
                                {
                                    sw.Write("<x:numFmt numFmtId=\"{0}\" formatCode=\"{1}\" />", listKeys[i], SLTool.XmlWrite(dictStyleNumberingFormat[listKeys[i]], gbThrowExceptionsIfAny));
                                }
                                sw.Write("</x:numFmts>");
                            }
                        }

                        sw.Write("<x:fonts count=\"{0}\">", listStyleFont.Count);
                        for (i = 0; i < listStyleFont.Count; ++i)
                        {
                            sw.Write("<x:font>{0}</x:font>", listStyleFont[i]);
                        }
                        sw.Write("</x:fonts>");

                        sw.Write("<x:fills count=\"{0}\">", listStyleFill.Count);
                        for (i = 0; i < listStyleFill.Count; ++i)
                        {
                            sw.Write("<x:fill>{0}</x:fill>", listStyleFill[i]);
                        }
                        sw.Write("</x:fills>");

                        List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();

                        sw.Write("<x:borders count=\"{0}\">", listStyleBorder.Count);
                        SLBorder slb;
                        for (i = 0; i < listStyleBorder.Count; ++i)
                        {
                            slb = new SLBorder(listempty, listempty);
                            slb.FromHash(listStyleBorder[i]);
                            slb.Sync();

                            sw.Write("<x:border");
                            if (slb.DiagonalUp != null) sw.Write(" diagonalUp=\"{0}\"", slb.DiagonalUp.Value ? "1" : "0");
                            if (slb.DiagonalDown != null) sw.Write(" diagonalDown=\"{0}\"", slb.DiagonalDown.Value ? "1" : "0");
                            if (slb.Outline != null && !slb.Outline.Value) sw.Write(" outline=\"0\"");
                            sw.Write(">");

                            // by "default" always have left, right, top, bottom and diagonal borders, even if empty?
                            sw.Write(SLBorderProperties.WriteToXmlTag("left", slb.LeftBorder));
                            sw.Write(SLBorderProperties.WriteToXmlTag("right", slb.RightBorder));
                            sw.Write(SLBorderProperties.WriteToXmlTag("top", slb.TopBorder));
                            sw.Write(SLBorderProperties.WriteToXmlTag("bottom", slb.BottomBorder));
                            sw.Write(SLBorderProperties.WriteToXmlTag("diagonal", slb.DiagonalBorder));
                            if (slb.HasVerticalBorder) sw.Write(SLBorderProperties.WriteToXmlTag("vertical", slb.VerticalBorder));
                            if (slb.HasHorizontalBorder) sw.Write(SLBorderProperties.WriteToXmlTag("horizontal", slb.HorizontalBorder));

                            sw.Write("</x:border>");
                        }
                        sw.Write("</x:borders>");

                        int iCanonicalCellStyleFormatId = 0;
                        SLStyle styleCanonical = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                        styleCanonical.FormatCode = this.NumberFormatGeneralText;
                        styleCanonical.fontReal.FromHash(listStyleFont[0]);
                        styleCanonical.Fill.SetPatternType(PatternValues.None);
                        styleCanonical.Border = new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                        iCanonicalCellStyleFormatId = this.SaveToStylesheetCellStylesFormat(styleCanonical.ToHash());
                        this.TranslateStylesToStyleIds(ref styleCanonical);

                        sw.Write("<x:cellStyleXfs count=\"{0}\">", listStyleCellStyleFormat.Count);
                        for (i = 0; i < listStyleCellStyleFormat.Count; ++i)
                        {
                            styleSL = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, listempty, listempty);
                            styleSL.FromHash(listStyleCellStyleFormat[i]);
                            this.TranslateStylesToStyleIds(ref styleSL);
                            sw.Write(styleSL.WriteToXmlTag());
                        }
                        sw.Write("</x:cellStyleXfs>");

                        sw.Write("<x:cellXfs count=\"{0}\">", listStyle.Count);
                        for (i = 0; i < listStyle.Count; ++i)
                        {
                            styleSL = new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, listempty, listempty);
                            styleSL.FromHash(listStyle[i]);
                            this.TranslateStylesToStyleIds(ref styleSL);
                            styleSL.CellStyleFormatId = (uint)iCanonicalCellStyleFormatId;
                            if (styleSL.NumberFormatId != styleCanonical.NumberFormatId) styleSL.ApplyNumberFormat = true;
                            if (styleSL.FontId != styleCanonical.FontId) styleSL.ApplyFont = true;
                            if (styleSL.FillId != styleCanonical.FillId) styleSL.ApplyFill = true;
                            if (styleSL.BorderId != styleCanonical.BorderId) styleSL.ApplyBorder = true;
                            if (styleSL.HasAlignment) styleSL.ApplyAlignment = true;
                            if (styleSL.HasProtection) styleSL.ApplyProtection = true;
                            sw.Write(styleSL.WriteToXmlTag());
                        }
                        sw.Write("</x:cellXfs>");

                        sw.Write("<x:cellStyles count=\"{0}\">", listStyleCellStyle.Count);
                        SLCellStyle csSL;
                        for (i = 0; i < listStyleCellStyle.Count; ++i)
                        {
                            csSL = new SLCellStyle();
                            csSL.FromHash(listStyleCellStyle[i]);
                            sw.Write(csSL.WriteToXmlTag());
                        }
                        sw.Write("</x:cellStyles>");

                        sw.Write("<x:dxfs count=\"{0}\">", listStyleDifferentialFormat.Count);
                        for (i = 0; i < listStyleDifferentialFormat.Count; ++i)
                        {
                            sw.Write("<x:dxf>{0}</x:dxf>", listStyleDifferentialFormat[i]);
                        }
                        sw.Write("</x:dxfs>");

                        sw.Write("<x:tableStyles count=\"{0}\"", listStyleTableStyle.Count);
                        if (this.TableStylesDefaultTableStyle.Length > 0)
                        {
                            sw.Write(" defaultTableStyle=\"{0}\"", this.TableStylesDefaultTableStyle);
                        }
                        if (this.TableStylesDefaultPivotStyle.Length > 0)
                        {
                            sw.Write(" defaultPivotStyle=\"{0}\"", this.TableStylesDefaultPivotStyle);
                        }
                        sw.Write(">");
                        SLTableStyle tsSL;
                        for (i = 0; i < listStyleTableStyle.Count; ++i)
                        {
                            tsSL = new SLTableStyle();
                            tsSL.FromHash(listStyleTableStyle[i]);
                            sw.Write(tsSL.WriteToXmlTag());
                        }
                        sw.Write("</x:tableStyles>");

                        sw.Write("</x:styleSheet>");

                        sw.Flush();
                        ms.Position = 0;
                        wbsp.FeedData(ms);
                    }
                }
                // end of writing new stylesheet
            }
        }
    }
}
