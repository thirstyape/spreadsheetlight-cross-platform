using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// Sort data by column.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be sorted, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be sorted, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="SortByColumnName">The column name of the column to be sorted by, such as "AA".</param>
        /// <param name="SortAscending">True to sort in ascending order. False to sort in descending order.</param>
        public void Sort(string StartCellReference, string EndCellReference, string SortByColumnName, bool SortAscending)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iSortByColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                && SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iSortByColumnIndex = SLTool.ToColumnIndex(SortByColumnName);
                this.Sort(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, true, iSortByColumnIndex, SortAscending);
            }
        }

        /// <summary>
        /// Sort data by row.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be sorted, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be sorted, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="SortByRowIndex">The row index of the row to be sorted by.</param>
        /// <param name="SortAscending">True to sort in ascending order. False to sort in descending order.</param>
        public void Sort(string StartCellReference, string EndCellReference, int SortByRowIndex, bool SortAscending)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                && SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                this.Sort(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, false, SortByRowIndex, SortAscending);
            }
        }

        /// <summary>
        /// Sort data by column.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <param name="SortByColumnIndex">The column index of the column to be sorted by.</param>
        /// <param name="SortAscending">True to sort in ascending order. False to sort in descending order.</param>
        public void Sort(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, int SortByColumnIndex, bool SortAscending)
        {
            this.Sort(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, true, SortByColumnIndex, SortAscending);
        }

        /// <summary>
        /// Sort data either by column or row.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <param name="SortByColumn">True to sort by column. False to sort by row.</param>
        /// <param name="SortByIndex">The row or column index of the row or column to be sorted by, depending on <paramref name="SortByColumn"/></param>
        /// <param name="SortAscending">True to sort in ascending order. False to sort in descending order.</param>
        public void Sort(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool SortByColumn, int SortByIndex, bool SortAscending)
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

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            // if the given index is out of the data range, then don't have to sort.
            if (SortByColumn)
            {
                if (SortByIndex < iStartColumnIndex || SortByIndex > iEndColumnIndex) return;
            }
            else
            {
                if (SortByIndex < iStartRowIndex || SortByIndex > iEndRowIndex) return;
            }

            SLCellWarehouse datacells = new SLCellWarehouse();
            int iRowIndex, iColumnIndex;
            int i, j;
            for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
            {
                for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                {
                    iRowIndex = i;
                    iColumnIndex = j;
                    if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                    {
                        datacells.SetValue(iRowIndex, iColumnIndex, slws.CellWarehouse.Cells[iRowIndex][iColumnIndex]);
                        slws.CellWarehouse.Remove(iRowIndex, iColumnIndex);
                    }
                }
            }

            List<SLSortItem> listNumbers = new List<SLSortItem>();
            List<SLSortItem> listText = new List<SLSortItem>();
            List<SLSortItem> listBoolean = new List<SLSortItem>();
            List<SLSortItem> listEmpty = new List<SLSortItem>();

            bool bValue = false;
            double fValue = 0.0;
            string sText = string.Empty;
            SLRstType rst;
            int index = 0;
            int iStartIndex = -1;
            int iEndIndex = -1;

            if (SortByColumn)
            {
                iStartIndex = iStartRowIndex;
                iEndIndex = iEndRowIndex;
            }
            else
            {
                iStartIndex = iStartColumnIndex;
                iEndIndex = iEndColumnIndex;
            }

            for (i = iStartIndex; i <= iEndIndex; ++i)
            {
                if (SortByColumn)
                {
                    iRowIndex = i;
                    iColumnIndex = SortByIndex;
                }
                else
                {
                    iRowIndex = SortByIndex;
                    iColumnIndex = i;
                }

                if (datacells.Exists(iRowIndex, iColumnIndex))
                {
                    if (datacells.Cells[iRowIndex][iColumnIndex].DataType == CellValues.Number)
                    {
                        if (datacells.Cells[iRowIndex][iColumnIndex].CellText != null)
                        {
                            if (double.TryParse(datacells.Cells[iRowIndex][iColumnIndex].CellText, out fValue))
                            {
                                listNumbers.Add(new SLSortItem() { Number = fValue, Index = i });
                            }
                            else
                            {
                                listText.Add(new SLSortItem() { Text = datacells.Cells[iRowIndex][iColumnIndex].CellText, Index = i });
                            }
                        }
                        else
                        {
                            listNumbers.Add(new SLSortItem() { Number = datacells.Cells[iRowIndex][iColumnIndex].NumericValue, Index = i });
                        }
                    }
                    else if (datacells.Cells[iRowIndex][iColumnIndex].DataType == CellValues.SharedString)
                    {
                        index = -1;

                        if (datacells.Cells[iRowIndex][iColumnIndex].CellText != null)
                        {
                            if (int.TryParse(datacells.Cells[iRowIndex][iColumnIndex].CellText, out index)
                                && index >= 0 && index < listSharedString.Count)
                            {
                                rst = new SLRstType(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                                rst.FromSharedStringItem(new SharedStringItem() { InnerXml = listSharedString[index] });
                                listText.Add(new SLSortItem() { Text = rst.ToPlainString(), Index = i });
                            }
                            else
                            {
                                listText.Add(new SLSortItem() { Text = datacells.Cells[iRowIndex][iColumnIndex].CellText, Index = i });
                            }
                        }
                        else
                        {
                            index = Convert.ToInt32(datacells.Cells[iRowIndex][iColumnIndex].NumericValue);
                            if (index >= 0 && index < listSharedString.Count)
                            {
                                rst = new SLRstType(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                                rst.FromSharedStringItem(new SharedStringItem() { InnerXml = listSharedString[index] });
                                listText.Add(new SLSortItem() { Text = rst.ToPlainString(), Index = i });
                            }
                            else
                            {
                                listText.Add(new SLSortItem() { Text = datacells.Cells[iRowIndex][iColumnIndex].NumericValue.ToString(CultureInfo.InvariantCulture), Index = i });
                            }
                        }
                    }
                    else if (datacells.Cells[iRowIndex][iColumnIndex].DataType == CellValues.Boolean)
                    {
                        if (datacells.Cells[iRowIndex][iColumnIndex].CellText != null)
                        {
                            if (double.TryParse(datacells.Cells[iRowIndex][iColumnIndex].CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out fValue))
                            {
                                listBoolean.Add(new SLSortItem() { Number = fValue > 0.5 ? 1.0 : 0.0, Index = i });
                            }
                            else if (bool.TryParse(datacells.Cells[iRowIndex][iColumnIndex].CellText, out bValue))
                            {
                                listBoolean.Add(new SLSortItem() { Number = bValue ? 1.0 : 0.0, Index = i });
                            }
                            else
                            {
                                listText.Add(new SLSortItem() { Text = datacells.Cells[iRowIndex][iColumnIndex].CellText, Index = i });
                            }
                        }
                        else
                        {
                            listBoolean.Add(new SLSortItem() { Number = datacells.Cells[iRowIndex][iColumnIndex].NumericValue > 0.5 ? 1.0 : 0.0, Index = i });
                        }
                    }
                    else
                    {
                        listText.Add(new SLSortItem() { Text = datacells.Cells[iRowIndex][iColumnIndex].CellText, Index = i });
                    }
                }
                else
                {
                    listEmpty.Add(new SLSortItem() { Index = i });
                }
            }

            listNumbers.Sort(new SLSortItemNumberComparer());
            if (!SortAscending) listNumbers.Reverse();

            listText.Sort(new SLSortItemTextComparer());
            if (!SortAscending) listText.Reverse();

            listBoolean.Sort(new SLSortItemNumberComparer());
            if (!SortAscending) listBoolean.Reverse();

            Dictionary<int, int> ReverseIndex = new Dictionary<int,int>();
            if (SortAscending)
            {
                j = iStartIndex;
                for (i = 0; i < listNumbers.Count; ++i)
                {
                    ReverseIndex[listNumbers[i].Index] = j;
                    ++j;
                }

                for (i = 0; i < listText.Count; ++i)
                {
                    ReverseIndex[listText[i].Index] = j;
                    ++j;
                }

                for (i = 0; i < listBoolean.Count; ++i)
                {
                    ReverseIndex[listBoolean[i].Index] = j;
                    ++j;
                }

                for (i = 0; i < listEmpty.Count; ++i)
                {
                    ReverseIndex[listEmpty[i].Index] = j;
                    ++j;
                }
            }
            else
            {
                j = iStartIndex;
                for (i = 0; i < listBoolean.Count; ++i)
                {
                    ReverseIndex[listBoolean[i].Index] = j;
                    ++j;
                }

                for (i = 0; i < listText.Count; ++i)
                {
                    ReverseIndex[listText[i].Index] = j;
                    ++j;
                }

                for (i = 0; i < listNumbers.Count; ++i)
                {
                    ReverseIndex[listNumbers[i].Index] = j;
                    ++j;
                }

                for (i = 0; i < listEmpty.Count; ++i)
                {
                    ReverseIndex[listEmpty[i].Index] = j;
                    ++j;
                }
            }

            List<int> listRowKeys = datacells.Cells.Keys.ToList<int>();
            List<int> listColumnKeys;
            int iNewRowIndex, iNewColumnIndex;
            foreach (int rowkey in listRowKeys)
            {
                listColumnKeys = datacells.Cells[rowkey].Keys.ToList<int>();
                foreach (int colkey in listColumnKeys)
                {
                    iRowIndex = rowkey;
                    iColumnIndex = colkey;
                    if (SortByColumn)
                    {
                        if (ReverseIndex.ContainsKey(iRowIndex))
                        {
                            iNewRowIndex = ReverseIndex[iRowIndex];
                            iNewColumnIndex = iColumnIndex;
                        }
                        else
                        {
                            // shouldn't happen, but just in case...
                            iNewRowIndex = iRowIndex;
                            iNewColumnIndex = iColumnIndex;
                        }
                    }
                    else
                    {
                        if (ReverseIndex.ContainsKey(iColumnIndex))
                        {
                            iNewRowIndex = iRowIndex;
                            iNewColumnIndex = ReverseIndex[iColumnIndex];
                        }
                        else
                        {
                            // shouldn't happen, but just in case...
                            iNewRowIndex = iRowIndex;
                            iNewColumnIndex = iColumnIndex;
                        }
                    }

                    slws.CellWarehouse.SetValue(iNewRowIndex, iNewColumnIndex, datacells.Cells[iRowIndex][iColumnIndex]);
                }
            }
        }
    }
}
