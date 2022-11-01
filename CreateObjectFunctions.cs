using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLC = SpreadsheetLight.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// Creates an instance of SLFont with theme information.
        /// </summary>
        /// <returns>An SLFont with theme information.</returns>
        public SLFont CreateFont()
        {
            return new SLFont(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
        }

        /// <summary>
        /// Creates an instance of SLPatternFill with theme information.
        /// </summary>
        /// <returns>An SLPatternFill with theme information.</returns>
        public SLPatternFill CreatePatternFill()
        {
            return new SLPatternFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
        }

        /// <summary>
        /// Creates an instance of SLGradientFill with theme information.
        /// </summary>
        /// <returns>An SLGradientFill with theme information.</returns>
        public SLGradientFill CreateGradientFill()
        {
            return new SLGradientFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
        }

        /// <summary>
        /// Creates an instance of SLFill with theme information.
        /// </summary>
        /// <returns>An SLFill with theme information.</returns>
        public SLFill CreateFill()
        {
            return new SLFill(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
        }

        /// <summary>
        /// Creates an instance of SLBorder with theme information.
        /// </summary>
        /// <returns>An SLBorder with theme information.</returns>
        public SLBorder CreateBorder()
        {
            return new SLBorder(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
        }

        /// <summary>
        /// Creates an instance of SLRstType with theme information.
        /// </summary>
        /// <returns>An SLRstType with theme information.</returns>
        public SLRstType CreateRstType()
        {
            return new SLRstType(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
        }

        /// <summary>
        /// Creates an instance of SLStyle with theme information.
        /// </summary>
        /// <returns>An SLStyle with theme information.</returns>
        public SLStyle CreateStyle()
        {
            return new SLStyle(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
        }

        /// <summary>
        /// Creates an instance of SLComment with theme information.
        /// </summary>
        /// <returns>An SLComment with theme information.</returns>
        public SLComment CreateComment()
        {
            SLComment comm = new SLComment(SimpleTheme.listThemeColors, SimpleTheme.ThrowExceptionsIfAny);
            if (this.DocumentProperties.Creator.Length > 0)
            {
                comm.Author = this.DocumentProperties.Creator;
            }
            else
            {
                comm.Author = SLConstants.ApplicationName;
            }

            return comm;
        }

        /// <summary>
        /// Creates an instance of SLDataValidation.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>An SLDataValidation.</returns>
        public SLDataValidation CreateDataValidation(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                iRowIndex = 1;
                iColumnIndex = 1;
            }

            SLDataValidation dv = new SLDataValidation();
            dv.InitialiseDataValidation(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex, slwb.WorkbookProperties.Date1904);
            return dv;
        }

        /// <summary>
        /// Creates an instance of SLDataValidation.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <returns>An SLDataValidation.</returns>
        public SLDataValidation CreateDataValidation(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iStartRowIndex = 1;
                iStartColumnIndex = 1;
                iEndRowIndex = 1;
                iEndColumnIndex = 1;
            }

            SLDataValidation dv = new SLDataValidation();
            dv.InitialiseDataValidation(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, slwb.WorkbookProperties.Date1904);
            return dv;
        }

        /// <summary>
        /// Creates an instance of SLDataValidation.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>An SLDataValidation.</returns>
        public SLDataValidation CreateDataValidation(int RowIndex, int ColumnIndex)
        {
            SLDataValidation dv = new SLDataValidation();
            dv.InitialiseDataValidation(RowIndex, ColumnIndex, RowIndex, ColumnIndex, slwb.WorkbookProperties.Date1904);
            return dv;
        }

        /// <summary>
        /// Creates an instance of SLDataValidation.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row.</param>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndRowIndex">The row index of the end row.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        /// <returns>An SLDataValidation.</returns>
        public SLDataValidation CreateDataValidation(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            SLDataValidation dv = new SLDataValidation();
            dv.InitialiseDataValidation(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, slwb.WorkbookProperties.Date1904);
            return dv;
        }

        /// <summary>
        /// Creates an instance of SLTable, given cell references of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be in the table, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be in the table, such as "A1". This is typically the bottom-right cell.</param>
        /// <returns>An SLTable with the required information.</returns>
        public SLTable CreateTable(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            return CreateTable(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Creates an instance of SLTable, given row and column indices of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <returns>An SLTable with the required information.</returns>
        public SLTable CreateTable(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
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
            if (iStartRowIndex == SLConstants.RowLimit) iStartRowIndex = SLConstants.RowLimit - 1;
            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            // consider minus 1 in case there's a totals row so there's less checking...
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            if (iEndRowIndex <= iStartRowIndex) iEndRowIndex = iStartRowIndex + 1;

            SLTable tbl = new SLTable();
            tbl.SetAllNull();

            slwb.RefreshPossibleTableId();
            tbl.Id = slwb.PossibleTableId;
            tbl.DisplayName = string.Format("Table{0}", tbl.Id);
            tbl.Name = tbl.DisplayName;

            tbl.StartRowIndex = iStartRowIndex;
            tbl.StartColumnIndex = iStartColumnIndex;
            tbl.EndRowIndex = iEndRowIndex;
            tbl.EndColumnIndex = iEndColumnIndex;

            tbl.AutoFilter.StartRowIndex = tbl.StartRowIndex;
            tbl.AutoFilter.StartColumnIndex = tbl.StartColumnIndex;
            tbl.AutoFilter.EndRowIndex = tbl.EndRowIndex;
            tbl.AutoFilter.EndColumnIndex = tbl.EndColumnIndex;
            tbl.HasAutoFilter = true;

            SLTableColumn tc;
            uint iColumnId = 1;
            int i, index;
            uint j;
            SLCell c;
            int iRowIndex, iColumnIndex;
            string sHeaderText = string.Empty;
            SharedStringItem ssi;
            SLRstType rst = new SLRstType(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            for (i = tbl.StartColumnIndex; i <= tbl.EndColumnIndex; ++i)
            {
                iRowIndex = StartRowIndex;
                iColumnIndex = i;
                sHeaderText = string.Empty;
                if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                {
                    c = slws.CellWarehouse.Cells[iRowIndex][iColumnIndex];
                    if (c.CellText == null)
                    {
                        if (c.DataType == CellValues.Number) sHeaderText = c.NumericValue.ToString(CultureInfo.InvariantCulture);
                        else if (c.DataType == CellValues.Boolean) sHeaderText = c.NumericValue > 0.5 ? "TRUE" : "FALSE";
                        else sHeaderText = string.Empty;
                    }
                    else
                    {
                        sHeaderText = c.CellText;
                    }

                    if (c.DataType == CellValues.SharedString)
                    {
                        index = -1;
                        if (c.CellText != null)
                        {
                            if (int.TryParse(c.CellText, out index))
                            {
                                index = -1;
                            }
                        }
                        else
                        {
                            index = Convert.ToInt32(c.NumericValue);
                        }

                        if (index >= 0 && index < listSharedString.Count)
                        {
                            ssi = new SharedStringItem();
                            ssi.InnerXml = listSharedString[index];
                            rst.FromSharedStringItem(ssi);
                            sHeaderText = rst.ToPlainString();
                        }
                    }
                }

                j = iColumnId;
                if (sHeaderText.Length == 0)
                {
                    sHeaderText = string.Format("Column{0}", j);
                }
                while (tbl.TableNames.Contains(sHeaderText))
                {
                    ++j;
                    sHeaderText = string.Format("Column{0}", j);
                }
                tc = new SLTableColumn();
                tc.Id = iColumnId;
                tc.Name = sHeaderText;
                tbl.TableColumns.Add(tc);
                tbl.TableNames.Add(sHeaderText);
                ++iColumnId;
            }

            tbl.TableStyleInfo.ShowFirstColumn = false;
            tbl.TableStyleInfo.ShowLastColumn = false;
            tbl.TableStyleInfo.ShowRowStripes = true;
            tbl.TableStyleInfo.ShowColumnStripes = false;
            tbl.HasTableStyleInfo = true;

            return tbl;
        }

        //Not ready yet!
        //public SLPivotTable CreatePivotTable(string StartCellReference, string EndCellReference)
        //{
        //    int iStartRowIndex = -1;
        //    int iStartColumnIndex = -1;
        //    int iEndRowIndex = -1;
        //    int iEndColumnIndex = -1;
        //    if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
        //    {
        //        iStartRowIndex = -1;
        //        iStartColumnIndex = -1;
        //    }
        //    if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
        //    {
        //        iEndRowIndex = -1;
        //        iEndColumnIndex = -1;
        //    }

        //    return CreatePivotTable(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        //}

        //public SLPivotTable CreatePivotTable(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        //{
        //    int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
        //    if (StartRowIndex < EndRowIndex)
        //    {
        //        iStartRowIndex = StartRowIndex;
        //        iEndRowIndex = EndRowIndex;
        //    }
        //    else
        //    {
        //        iStartRowIndex = EndRowIndex;
        //        iEndRowIndex = StartRowIndex;
        //    }

        //    if (StartColumnIndex < EndColumnIndex)
        //    {
        //        iStartColumnIndex = StartColumnIndex;
        //        iEndColumnIndex = EndColumnIndex;
        //    }
        //    else
        //    {
        //        iStartColumnIndex = EndColumnIndex;
        //        iEndColumnIndex = StartColumnIndex;
        //    }

        //    // not checking bounds because we're going to be more stringent on the data source range.

        //    SLPivotTable pivot = new SLPivotTable();
        //    slwb.RefreshPossiblePivotTableCacheId();
        //    pivot.CacheId = slwb.PossiblePivotTableCacheId;
        //    pivot.Name = slwb.GetNextPossiblePivotTableName();

        //    pivot.IsDataSourceTable = false;
        //    pivot.SheetTableName = gsSelectedWorksheetName;

        //    return pivot;
        //}

        /// <summary>
        /// Creates an instance of SLSparklineGroup, given cell references of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be in the sparkline, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be in the sparkline, such as "A1". This is typically the bottom-right cell.</param>
        /// <returns>An SLSparklineGroup with the required information.</returns>
        public SLSparklineGroup CreateSparklineGroup(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            return this.CreateSparklineGroup(gsSelectedWorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Creates an instance of SLSparklineGroup, given cell references of opposite cells in a cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be in the sparkline, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be in the sparkline, such as "A1". This is typically the bottom-right cell.</param>
        /// <returns>An SLSparklineGroup with the required information.</returns>
        public SLSparklineGroup CreateSparklineGroup(string WorksheetName, string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            return this.CreateSparklineGroup(WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Creates an instance of SLSparklineGroup, given row and column indices of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <returns>An SLSparklineGroup with the required information.</returns>
        public SLSparklineGroup CreateSparklineGroup(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return this.CreateSparklineGroup(gsSelectedWorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex);
        }

        /// <summary>
        /// Creates an instance of SLSparklineGroup, given row and column indices of opposite cells in a cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <returns>An SLSparklineGroup with the required information.</returns>
        public SLSparklineGroup CreateSparklineGroup(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            SLSparklineGroup spk = new SLSparklineGroup(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);

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

            spk.WorksheetName = WorksheetName;
            spk.StartRowIndex = iStartRowIndex;
            spk.StartColumnIndex = iStartColumnIndex;
            spk.EndRowIndex = iEndRowIndex;
            spk.EndColumnIndex = iEndColumnIndex;

            // this seems to be the "default"
            spk.SetSparklineStyle(SLSparklineStyle.Accent1Darker50Percent);

            return spk;
        }

        /// <summary>
        /// Creates an instance of SLChart, given cell references of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be in the chart, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be in the chart, such as "A1". This is typically the bottom-right cell.</param>
        /// <returns>An SLChart with the required information.</returns>
        public SLC.SLChart CreateChart(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            return this.CreateChartInternal(gsSelectedWorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, new SLC.SLCreateChartOptions());
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Creates an instance of SLChart, given cell references of opposite cells in a cell range and whether rows or columns contain the data series and if hidden data is used.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be in the chart, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be in the chart, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="RowsAsDataSeries">True if rows contain the data series. False if columns contain the data series.</param>
        /// <param name="ShowHiddenData">True if hidden data is used in the chart. False otherwise.</param>
        /// <returns>An SLChart with the required information.</returns>
        [Obsolete("Use an overload with the SLCreateChartOptions parameter instead.")]
        public SLC.SLChart CreateChart(string StartCellReference, string EndCellReference, bool RowsAsDataSeries, bool ShowHiddenData)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            SLC.SLCreateChartOptions Options = new SLC.SLCreateChartOptions();
            Options.RowsAsDataSeries = RowsAsDataSeries;
            Options.ShowHiddenData = ShowHiddenData;
            return this.CreateChartInternal(gsSelectedWorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Options);
        }

        /// <summary>
        /// Creates an instance of SLChart, given cell references of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be in the chart, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be in the chart, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="Options">Chart creation options.</param>
        /// <returns>An SLChart with the required information.</returns>
        public SLC.SLChart CreateChart(string StartCellReference, string EndCellReference, SLC.SLCreateChartOptions Options)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            return this.CreateChartInternal(gsSelectedWorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Options);
        }

        /// <summary>
        /// Creates an instance of SLChart, given cell references of opposite cells in a cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be in the chart, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be in the chart, such as "A1". This is typically the bottom-right cell.</param>
        /// <returns>An SLChart with the required information.</returns>
        public SLC.SLChart CreateChart(string WorksheetName, string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            return this.CreateChartInternal(WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, new SLC.SLCreateChartOptions());
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Creates an instance of SLChart, given cell references of opposite cells in a cell range and whether rows or columns contain the data series.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be in the chart, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be in the chart, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="RowsAsDataSeries">True if rows contain the data series. False if columns contain the data series.</param>
        /// <param name="ShowHiddenData">True if hidden data is used in the chart. False otherwise.</param>
        /// <returns>An SLChart with the required information.</returns>
        [Obsolete("Use an overload with the SLCreateChartOptions parameter instead.")]
        public SLC.SLChart CreateChart(string WorksheetName, string StartCellReference, string EndCellReference, bool RowsAsDataSeries, bool ShowHiddenData)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            SLC.SLCreateChartOptions Options = new SLC.SLCreateChartOptions();
            Options.RowsAsDataSeries = RowsAsDataSeries;
            Options.ShowHiddenData = ShowHiddenData;
            return this.CreateChartInternal(WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Options);
        }

        /// <summary>
        /// Creates an instance of SLChart, given cell references of opposite cells in a cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be in the chart, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be in the chart, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="Options">Chart creation options.</param>
        /// <returns>An SLChart with the required information.</returns>
        public SLC.SLChart CreateChart(string WorksheetName, string StartCellReference, string EndCellReference, SLC.SLCreateChartOptions Options)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            return this.CreateChartInternal(WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Options);
        }

        /// <summary>
        /// Creates an instance of SLChart, given row and column indices of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <returns>An SLChart with the required information.</returns>
        public SLC.SLChart CreateChart(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return this.CreateChartInternal(gsSelectedWorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, new SLC.SLCreateChartOptions());
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Creates an instance of SLChart, given row and column indices of opposite cells in a cell range and whether rows or columns contain the data series and if hidden data is used.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <param name="RowsAsDataSeries">True if rows contain the data series. False if columns contain the data series.</param>
        /// <param name="ShowHiddenData">True if hidden data is used in the chart. False otherwise.</param>
        /// <returns>An SLChart with the required information.</returns>
        [Obsolete("Use an overload with the SLCreateChartOptions parameter instead.")]
        public SLC.SLChart CreateChart(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool RowsAsDataSeries, bool ShowHiddenData)
        {
            SLC.SLCreateChartOptions Options = new SLC.SLCreateChartOptions();
            Options.RowsAsDataSeries = RowsAsDataSeries;
            Options.ShowHiddenData = ShowHiddenData;
            return this.CreateChartInternal(gsSelectedWorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, Options);
        }

        /// <summary>
        /// Creates an instance of SLChart, given row and column indices of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <param name="Options">Chart creation options.</param>
        /// <returns>An SLChart with the required information.</returns>
        public SLC.SLChart CreateChart(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLC.SLCreateChartOptions Options)
        {
            return this.CreateChartInternal(gsSelectedWorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, Options);
        }

        /// <summary>
        /// Creates an instance of SLChart, given row and column indices of opposite cells in a cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <returns>An SLChart with the required information.</returns>
        public SLC.SLChart CreateChart(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return this.CreateChartInternal(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, new SLC.SLCreateChartOptions());
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Creates an instance of SLChart, given row and column indices of opposite cells in a cell range and whether rows or columns contain the data series and if hidden data is used.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <param name="RowsAsDataSeries">True if rows contain the data series. False if columns contain the data series.</param>
        /// <param name="ShowHiddenData">True if hidden data is used in the chart. False otherwise.</param>
        /// <returns>An SLChart with the required information.</returns>
        [Obsolete("Use an overload with the SLCreateChartOptions parameter instead.")]
        public SLC.SLChart CreateChart(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool RowsAsDataSeries, bool ShowHiddenData)
        {
            SLC.SLCreateChartOptions Options = new SLC.SLCreateChartOptions();
            Options.RowsAsDataSeries = RowsAsDataSeries;
            Options.ShowHiddenData = ShowHiddenData;
            return this.CreateChartInternal(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, Options);
        }

        /// <summary>
        /// Creates an instance of SLChart, given row and column indices of opposite cells in a cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <param name="Options">Chart creation options.</param>
        /// <returns>An SLChart with the required information.</returns>
        public SLC.SLChart CreateChart(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLC.SLCreateChartOptions Options)
        {
            return this.CreateChartInternal(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, Options);
        }

        private SLC.SLChart CreateChartInternal(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLC.SLCreateChartOptions Options)
        {
            if (Options == null) Options = new SLC.SLCreateChartOptions();

            SLC.SLChart chart = new SLC.SLChart();

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

            // this will keep the calculations within workable range
            if (iStartRowIndex >= SLConstants.RowLimit) iStartRowIndex = SLConstants.RowLimit - 1;
            if (iStartColumnIndex >= SLConstants.ColumnLimit) iStartColumnIndex = SLConstants.ColumnLimit - 1;

            chart.WorksheetName = WorksheetName;

            if (Options.RowsAsDataSeries == null)
            {
                if ((iEndColumnIndex - iStartColumnIndex) >= (iEndRowIndex - iStartRowIndex))
                {
                    chart.RowsAsDataSeries = true;
                }
                else
                {
                    chart.RowsAsDataSeries = false;
                }
            }
            else
            {
                chart.RowsAsDataSeries = Options.RowsAsDataSeries.Value;
            }

            chart.ShowHiddenData = Options.ShowHiddenData;
            chart.ShowDataLabelsOverMaximum = Options.IsStylish ? false : true;

            int i;
            chart.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < SimpleTheme.listThemeColors.Count; ++i)
            {
                chart.listThemeColors.Add(SimpleTheme.listThemeColors[i]);
            }

            chart.Date1904 = this.slwb.WorkbookProperties.Date1904;
            chart.IsStylish = Options.IsStylish;
            chart.RoundedCorners = false;

            chart.ThrowExceptionsIfAny = SimpleTheme.ThrowExceptionsIfAny;

            // assume combination charts are possible first
            chart.IsCombinable = true;

            chart.PlotArea = new SLC.SLPlotArea(SimpleTheme.listThemeColors, slwb.WorkbookProperties.Date1904, Options.IsStylish, SimpleTheme.ThrowExceptionsIfAny);
            chart.PlotArea.DataSeries = this.FillChartDataSeries(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, chart.RowsAsDataSeries, chart.ShowHiddenData);
            chart.SetPlotAreaAxes();

            chart.HasShownSecondaryTextAxis = false;

            chart.StartRowIndex = iStartRowIndex;
            chart.StartColumnIndex = iStartColumnIndex;
            chart.EndRowIndex = iEndRowIndex;
            chart.EndColumnIndex = iEndColumnIndex;

            chart.ShowEmptyCellsAs = Options.IsStylish ? C.DisplayBlanksAsValues.Zero : C.DisplayBlanksAsValues.Gap;
            chart.ChartStyle = SLC.SLChartStyle.Style2;

            chart.ChartName = string.Format("Chart {0}", slws.Charts.Count + 1);

            chart.HasTitle = false;
            chart.Title = new SLC.SLTitle(SimpleTheme.listThemeColors, Options.IsStylish, SimpleTheme.ThrowExceptionsIfAny);
            chart.Title.Overlay = false;

            chart.Is3D = false;
            chart.Floor = new SLC.SLFloor(SimpleTheme.listThemeColors, Options.IsStylish, SimpleTheme.ThrowExceptionsIfAny);
            chart.SideWall = new SLC.SLSideWall(SimpleTheme.listThemeColors, Options.IsStylish, SimpleTheme.ThrowExceptionsIfAny);
            chart.BackWall = new SLC.SLBackWall(SimpleTheme.listThemeColors, Options.IsStylish, SimpleTheme.ThrowExceptionsIfAny);

            chart.ShowLegend = true;
            chart.Legend = new SLC.SLLegend(SimpleTheme.listThemeColors, Options.IsStylish, SimpleTheme.ThrowExceptionsIfAny);
            chart.Legend.Overlay = false;
            if (Options.IsStylish)
            {
                chart.Legend.LegendPosition = A.Charts.LegendPositionValues.Bottom;
            }

            chart.ShapeProperties = new SLA.SLShapeProperties(SimpleTheme.listThemeColors, SimpleTheme.ThrowExceptionsIfAny);

            if (Options.IsStylish)
            {
                chart.ShapeProperties.Fill.SetSolidFill(A.SchemeColorValues.Background1, 0, 0);
                chart.ShapeProperties.Outline.Width = 0.75m;
                chart.ShapeProperties.Outline.CapType = A.LineCapValues.Flat;
                chart.ShapeProperties.Outline.CompoundLineType = A.CompoundLineValues.Single;
                chart.ShapeProperties.Outline.Alignment = A.PenAlignmentValues.Center;
                chart.ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                chart.ShapeProperties.Outline.JoinType = SLA.SLLineJoinValues.Round;
            }

            chart.TopPosition = 1;
            chart.LeftPosition = 1;
            chart.BottomPosition = 16;
            chart.RightPosition = 8.5;

            return chart;
        }

        // this is here because it's only used by the FillChartDataSeries() function.
        private string GetCellTrueValue(SLCell c)
        {
            string sValue = c.CellText ?? string.Empty;
            if (c.DataType == CellValues.Number)
            {
                if (c.CellText == null)
                {
                    // apparently we can only print up to a limited number of decimal places,
                    // albeit a large number. This is a limitation on the double variable type.
                    // Excel can print more decimal places. You go Excel...
                    // Go Google IEEE and the floating point standard for more details.
                    
                    // We could store using a decimal type in SLCell, but I don't think
                    // it's worth it given speed vs accuracy vs number range issues.
                    // If you need larger number of decimal places of accuracy in a chart,
                    // you've probably got a problem... No one's gonna be able to tell if
                    // there's a difference anyway... And if you need to present that many
                    // decimal places of accuracy, a chart is probably the wrong method of
                    // displaying it.
                    // You really need the extra decimal places? Try setting the original values
                    // with SetCellValueNumeric() and use up to the desired accuracy in the string.
                    sValue = c.NumericValue.ToString(CultureInfo.InvariantCulture);
                }
            }
            else if (c.DataType == CellValues.SharedString)
            {
                if (c.CellText == null)
                {
                    int index = Convert.ToInt32(c.NumericValue);
                    SLRstType rst;
                    sValue = string.Empty;
                    if (index >= 0 && index < listSharedString.Count)
                    {
                        rst = new SLRstType(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                        rst.FromSharedStringItem(new SharedStringItem() { InnerXml = listSharedString[index] });
                        sValue = rst.ToPlainString();
                    }
                }
            }
            else if (c.DataType == CellValues.Boolean)
            {
                if (c.CellText == null)
                {
                    if (c.NumericValue > 0.5) sValue = "1";
                    else sValue = "0";
                }
            }
            // no inline strings

            return sValue;
        }

        private List<SLC.SLDataSeries> FillChartDataSeries(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool RowsAsDataSeries, bool ShowHiddenData)
        {
            List<SLC.SLDataSeries> series = new List<SLC.SLDataSeries>();
            int i, j;
            SLCell c;
            int iRowIndex, iColumnIndex;
            Dictionary<int, bool> HiddenRows = new Dictionary<int, bool>();
            Dictionary<int, bool> HiddenColumns = new Dictionary<int, bool>();
            SLCellWarehouse cellstore = new SLCellWarehouse();

            #region GetCells
            for (i = StartRowIndex; i <= EndRowIndex; ++i)
            {
                HiddenRows[i] = false;
            }
            for (j = StartColumnIndex; j <= EndColumnIndex; ++j)
            {
                HiddenColumns[j] = false;
            }

            bool bFound = false;
            string sWorksheetRelID = string.Empty;
            if (WorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                bFound = false;
            }
            else
            {
                bFound = false;
                foreach (SLSheet sheet in slwb.Sheets)
                {
                    if (sheet.Name.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        bFound = true;
                        sWorksheetRelID = sheet.Id;
                        break;
                    }
                }
            }

            if (bFound)
            {
                Dictionary<int, HashSet<int>> multiCellRef = new Dictionary<int, HashSet<int>>();
                for (i = StartRowIndex; i <= EndRowIndex; ++i)
                {
                    multiCellRef.Add(i, new HashSet<int>());
                    for (j = StartColumnIndex; j <= EndColumnIndex; ++j)
                    {
                        multiCellRef[i].Add(j);
                    }
                }

                WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(sWorksheetRelID);
                Row row;
                Column column;
                Cell cell;
                if (!ShowHiddenData)
                {
                    // running it twice because Row contains Cell classes and it's easier this way
                    using (OpenXmlReader oxr = OpenXmlReader.Create(wsp))
                    {
                        while (oxr.Read())
                        {
                            if (oxr.ElementType == typeof(Row))
                            {
                                row = (Row)oxr.LoadCurrentElement();
                                if (row.RowIndex != null)
                                {
                                    foreach (var rowindex in HiddenRows.Keys)
                                    {
                                        if (row.RowIndex.Value == rowindex && row.Hidden != null && row.Hidden.Value)
                                        {
                                            HiddenRows[rowindex] = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                using (OpenXmlReader oxr = OpenXmlReader.Create(wsp))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(Column))
                        {
                            if (!ShowHiddenData)
                            {
                                column = (Column)oxr.LoadCurrentElement();
                                foreach (var colindex in HiddenColumns.Keys)
                                {
                                    if (column.Min <= colindex && colindex <= column.Max && column.Hidden != null && column.Hidden.Value)
                                    {
                                        HiddenColumns[colindex] = true;
                                    }
                                }
                            }
                        }
                        else if (oxr.ElementType == typeof(Cell))
                        {
                            cell = (Cell)oxr.LoadCurrentElement();
                            if (cell.CellReference != null)
                            {
                                if (SLTool.FormatCellReferenceToRowColumnIndex(cell.CellReference.Value, out i, out j))
                                {
                                    c = new SLCell();
                                    c.FromCell(cell);
                                    cellstore.SetValue(i, j, c);
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                SLRowProperties rp;
                SLColumnProperties cp;

                if (!ShowHiddenData)
                {
                    for (j = StartColumnIndex; j <= EndColumnIndex; ++j)
                    {
                        if (slws.ColumnProperties.ContainsKey(j))
                        {
                            cp = slws.ColumnProperties[j];
                            if (cp.Hidden) HiddenColumns[j] = true;
                        }
                    }
                }

                for (i = StartRowIndex; i <= EndRowIndex; ++i)
                {
                    if (!ShowHiddenData && slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        if (rp.Hidden) HiddenRows[i] = true;
                    }

                    for (j = StartColumnIndex; j <= EndColumnIndex; ++j)
                    {
                        if (slws.CellWarehouse.Exists(i, j))
                        {
                            cellstore.SetValue(i, j, slws.CellWarehouse.Cells[i][j]);
                        }
                    }
                }
            }

            #endregion

            int index = 0;
            int index2 = 0;
            string sCellValue;
            string sFormatCode;
            SLC.SLDataSeries ser;
            SLC.SLStringReference sr;
            SLC.SLNumberReference nr;
            SLStyle style;

            bool bIsStringReference = true;
            // we're going to assume that the format code applies to all in the "category" cells.
            string sAxisFormatCode = string.Empty;
            
            SLC.SLAxisDataSourceType cat = new SLC.SLAxisDataSourceType();
            if (RowsAsDataSeries)
            {
                bIsStringReference = true;
                sAxisFormatCode = SLConstants.NumberFormatGeneral;
                iRowIndex = StartRowIndex;
                iColumnIndex = StartColumnIndex + 1;
                if (cellstore.Exists(iRowIndex, iColumnIndex))
                {
                    // dates are also numbers, so we lump it together
                    if (cellstore.Cells[iRowIndex][iColumnIndex].DataType == CellValues.Number)
                    {
                        bIsStringReference = false;
                        style = new SLStyle(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                        style.FromHash(listStyle[(int)cellstore.Cells[iRowIndex][iColumnIndex].StyleIndex]);
                        this.TranslateStylesToStyleIds(ref style);
                        sAxisFormatCode = style.FormatCode;
                    }
                }

                sr = new SLC.SLStringReference();
                nr = new SLC.SLNumberReference();
                if (bIsStringReference)
                {
                    cat.UseStringReference = true;
                    sr = new SLC.SLStringReference();
                    sr.WorksheetName = WorksheetName;
                    sr.StartRowIndex = StartRowIndex;
                    sr.StartColumnIndex = StartColumnIndex + 1;
                    sr.EndRowIndex = StartRowIndex;
                    sr.EndColumnIndex = EndColumnIndex;
                    sr.Formula = SLC.SLChartTool.GetChartReferenceFormula(WorksheetName, StartRowIndex, StartColumnIndex + 1, StartRowIndex, EndColumnIndex);
                }
                else
                {
                    cat.UseNumberReference = true;
                    nr.WorksheetName = WorksheetName;
                    nr.StartRowIndex = StartRowIndex;
                    nr.StartColumnIndex = StartColumnIndex + 1;
                    nr.EndRowIndex = StartRowIndex;
                    nr.EndColumnIndex = EndColumnIndex;
                    nr.Formula = SLC.SLChartTool.GetChartReferenceFormula(WorksheetName, StartRowIndex, StartColumnIndex + 1, StartRowIndex, EndColumnIndex);
                    nr.NumberingCache.FormatCode = sAxisFormatCode;
                }

                index2 = 0;
                // if the header row is hidden, I don't know what to do...
                for (j = StartColumnIndex + 1; j <= EndColumnIndex; ++j)
                {
                    if (HiddenColumns.ContainsKey(j) && !HiddenColumns[j])
                    {
                        iRowIndex = StartRowIndex;
                        iColumnIndex = j;
                        sCellValue = string.Empty;
                        if (cellstore.Exists(iRowIndex, iColumnIndex))
                        {
                            c = cellstore.Cells[iRowIndex][iColumnIndex];
                            sCellValue = this.GetCellTrueValue(c);

                            if (bIsStringReference)
                            {
                                sr.Points.Add(new SLC.SLStringPoint()
                                {
                                    Index = (uint)index2,
                                    NumericValue = sCellValue
                                });
                            }
                            else
                            {
                                nr.NumberingCache.Points.Add(new SLC.SLNumericPoint()
                                {
                                    Index = (uint)index2,
                                    NumericValue = sCellValue
                                });
                            }
                        }
                        ++index2;
                    }
                }

                if (bIsStringReference)
                {
                    sr.PointCount = (uint)index2;
                    cat.StringReference = sr;
                }
                else
                {
                    nr.NumberingCache.PointCount = (uint)index2;
                    cat.NumberReference = nr;
                }

                index = 0;
                for (i = StartRowIndex + 1; i <= EndRowIndex; ++i)
                {
                    if (HiddenRows.ContainsKey(i) && !HiddenRows[i])
                    {
                        ser = new SLC.SLDataSeries(SimpleTheme.listThemeColors, SimpleTheme.ThrowExceptionsIfAny);
                        ser.Index = (uint)index;
                        ser.Order = (uint)index;
                        ser.IsStringReference = true;

                        sr = new SLC.SLStringReference();
                        iRowIndex = i;
                        iColumnIndex = StartColumnIndex;
                        sCellValue = string.Empty;
                        if (cellstore.Exists(iRowIndex, iColumnIndex))
                        {
                            c = cellstore.Cells[iRowIndex][iColumnIndex];
                            sCellValue = this.GetCellTrueValue(c);
                        }
                        sr.WorksheetName = WorksheetName;
                        sr.StartRowIndex = i;
                        sr.StartColumnIndex = StartColumnIndex;
                        sr.EndRowIndex = i;
                        sr.EndColumnIndex = StartColumnIndex;
                        sr.Formula = SLC.SLChartTool.GetChartReferenceFormula(WorksheetName, i, StartColumnIndex);
                        sr.PointCount = 1;
                        sr.Points.Add(new SLC.SLStringPoint() { Index = 0, NumericValue = sCellValue });
                        ser.StringReference = sr;

                        ser.AxisData = cat.Clone();

                        ser.NumberData.UseNumberReference = true;

                        nr = new SLC.SLNumberReference();
                        nr.WorksheetName = WorksheetName;
                        nr.StartRowIndex = i;
                        nr.StartColumnIndex = StartColumnIndex + 1;
                        nr.EndRowIndex = i;
                        nr.EndColumnIndex = EndColumnIndex;
                        nr.Formula = SLC.SLChartTool.GetChartReferenceFormula(WorksheetName, i, StartColumnIndex + 1, i, EndColumnIndex);
                        nr.NumberingCache.FormatCode = SLConstants.NumberFormatGeneral;
                        index2 = 0;
                        for (j = StartColumnIndex + 1; j <= EndColumnIndex; ++j)
                        {
                            if (HiddenColumns.ContainsKey(j) && !HiddenColumns[j])
                            {
                                iRowIndex = i;
                                iColumnIndex = j;
                                sCellValue = string.Empty;
                                sFormatCode = string.Empty;
                                if (cellstore.Exists(iRowIndex, iColumnIndex))
                                {
                                    c = cellstore.Cells[iRowIndex][iColumnIndex];
                                    sCellValue = this.GetCellTrueValue(c);

                                    style = new SLStyle(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                                    style.FromHash(listStyle[(int)c.StyleIndex]);
                                    this.TranslateStylesToStyleIds(ref style);
                                    if (style.HasNumberingFormat) sFormatCode = style.FormatCode;

                                    nr.NumberingCache.Points.Add(new SLC.SLNumericPoint()
                                    {
                                        FormatCode = sFormatCode,
                                        Index = (uint)index2,
                                        NumericValue = sCellValue
                                    });
                                }
                                ++index2;
                            }
                        }
                        nr.NumberingCache.PointCount = (uint)index2;
                        ser.NumberData.NumberReference = nr;

                        series.Add(ser);
                        ++index;
                    }
                }

                // end of rows as data series
            }
            else
            {
                bIsStringReference = true;
                sAxisFormatCode = SLConstants.NumberFormatGeneral;
                iRowIndex = StartRowIndex + 1;
                iColumnIndex = StartColumnIndex;
                if (cellstore.Exists(iRowIndex, iColumnIndex))
                {
                    // dates are also numbers, so we lump it together
                    if (cellstore.Cells[iRowIndex][iColumnIndex].DataType == CellValues.Number)
                    {
                        bIsStringReference = false;
                        style = new SLStyle(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                        style.FromHash(listStyle[(int)cellstore.Cells[iRowIndex][iColumnIndex].StyleIndex]);
                        this.TranslateStylesToStyleIds(ref style);
                        sAxisFormatCode = style.FormatCode;
                    }
                }

                sr = new SLC.SLStringReference();
                nr = new SLC.SLNumberReference();
                if (bIsStringReference)
                {
                    cat.UseStringReference = true;
                    sr.WorksheetName = WorksheetName;
                    sr.StartRowIndex = StartRowIndex + 1;
                    sr.StartColumnIndex = StartColumnIndex;
                    sr.EndRowIndex = EndRowIndex;
                    sr.EndColumnIndex = StartColumnIndex;
                    sr.Formula = SLC.SLChartTool.GetChartReferenceFormula(WorksheetName, StartRowIndex + 1, StartColumnIndex, EndRowIndex, StartColumnIndex);
                }
                else
                {
                    cat.UseNumberReference = true;
                    nr.WorksheetName = WorksheetName;
                    nr.StartRowIndex = StartRowIndex + 1;
                    nr.StartColumnIndex = StartColumnIndex;
                    nr.EndRowIndex = EndRowIndex;
                    nr.EndColumnIndex = StartColumnIndex;
                    nr.Formula = SLC.SLChartTool.GetChartReferenceFormula(WorksheetName, StartRowIndex + 1, StartColumnIndex, EndRowIndex, StartColumnIndex);
                    nr.NumberingCache.FormatCode = sAxisFormatCode;
                }
                
                index2 = 0;
                // if the header column is hidden, I don't know what to do...
                for (i = StartRowIndex + 1; i <= EndRowIndex; ++i)
                {
                    if (HiddenRows.ContainsKey(i) && !HiddenRows[i])
                    {
                        iRowIndex = i;
                        iColumnIndex = StartColumnIndex;
                        sCellValue = string.Empty;
                        if (cellstore.Exists(iRowIndex, iColumnIndex))
                        {
                            c = cellstore.Cells[iRowIndex][iColumnIndex];
                            sCellValue = this.GetCellTrueValue(c);

                            if (bIsStringReference)
                            {
                                sr.Points.Add(new SLC.SLStringPoint()
                                {
                                    Index = (uint)index2,
                                    NumericValue = sCellValue
                                });
                            }
                            else
                            {
                                nr.NumberingCache.Points.Add(new SLC.SLNumericPoint()
                                {
                                    Index = (uint)index2,
                                    NumericValue = sCellValue
                                });
                            }
                        }
                        ++index2;
                    }
                }

                if (bIsStringReference)
                {
                    sr.PointCount = (uint)index2;
                    cat.StringReference = sr;
                }
                else
                {
                    nr.NumberingCache.PointCount = (uint)index2;
                    cat.NumberReference = nr;
                }

                index = 0;
                for (j = StartColumnIndex + 1; j <= EndColumnIndex; ++j)
                {
                    if (HiddenColumns.ContainsKey(j) && !HiddenColumns[j])
                    {
                        ser = new SLC.SLDataSeries(SimpleTheme.listThemeColors, SimpleTheme.ThrowExceptionsIfAny);
                        ser.Index = (uint)index;
                        ser.Order = (uint)index;
                        ser.IsStringReference = true;

                        sr = new SLC.SLStringReference();
                        iRowIndex = StartRowIndex;
                        iColumnIndex = j;
                        sCellValue = string.Empty;
                        if (cellstore.Exists(iRowIndex, iColumnIndex))
                        {
                            c = cellstore.Cells[iRowIndex][iColumnIndex];
                            sCellValue = this.GetCellTrueValue(c);
                        }
                        sr.WorksheetName = WorksheetName;
                        sr.StartRowIndex = StartRowIndex;
                        sr.StartColumnIndex = j;
                        sr.EndRowIndex = StartRowIndex;
                        sr.EndColumnIndex = j;
                        sr.Formula = SLC.SLChartTool.GetChartReferenceFormula(WorksheetName, StartRowIndex, j);
                        sr.PointCount = 1;
                        sr.Points.Add(new SLC.SLStringPoint() { Index = 0, NumericValue = sCellValue });
                        ser.StringReference = sr;

                        ser.AxisData = cat.Clone();

                        ser.NumberData.UseNumberReference = true;

                        nr = new SLC.SLNumberReference();
                        nr.WorksheetName = WorksheetName;
                        nr.StartRowIndex = StartRowIndex + 1;
                        nr.StartColumnIndex = j;
                        nr.EndRowIndex = EndRowIndex;
                        nr.EndColumnIndex = j;
                        nr.Formula = SLC.SLChartTool.GetChartReferenceFormula(WorksheetName, StartRowIndex + 1, j, EndRowIndex, j);
                        nr.NumberingCache.FormatCode = SLConstants.NumberFormatGeneral;
                        index2 = 0;
                        for (i = StartRowIndex + 1; i <= EndRowIndex; ++i)
                        {
                            if (HiddenRows.ContainsKey(i) && !HiddenRows[i])
                            {
                                iRowIndex = i;
                                iColumnIndex = j;
                                sCellValue = string.Empty;
                                sFormatCode = string.Empty;
                                if (cellstore.Exists(iRowIndex, iColumnIndex))
                                {
                                    c = cellstore.Cells[iRowIndex][iColumnIndex];
                                    sCellValue = this.GetCellTrueValue(c);

                                    style = new SLStyle(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                                    style.FromHash(listStyle[(int)c.StyleIndex]);
                                    this.TranslateStylesToStyleIds(ref style);
                                    if (style.HasNumberingFormat) sFormatCode = style.FormatCode;

                                    nr.NumberingCache.Points.Add(new SLC.SLNumericPoint()
                                    {
                                        FormatCode = sFormatCode,
                                        Index = (uint)index2,
                                        NumericValue = sCellValue
                                    });
                                }
                                ++index2;
                            }
                        }
                        nr.NumberingCache.PointCount = (uint)index2;
                        ser.NumberData.NumberReference = nr;

                        series.Add(ser);
                        ++index;
                    }
                }

                // end of columns as data series
            }

            return series;
        }
    }
}
