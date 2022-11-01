using System;
using System.Collections.Generic;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Excel = DocumentFormat.OpenXml.Office.Excel;

namespace SpreadsheetLight
{
    /// <summary>
    /// Built-in sparkline styles.
    /// </summary>
    public enum SLSparklineStyle
    {
        /// <summary>
        /// Accent 1 Darker 50%
        /// </summary>
        Accent1Darker50Percent = 0,
        /// <summary>
        /// Accent 2 Darker 50%
        /// </summary>
        Accent2Darker50Percent,
        /// <summary>
        /// Accent 3 Darker 50%
        /// </summary>
        Accent3Darker50Percent,
        /// <summary>
        /// Accent 4 Darker 50%
        /// </summary>
        Accent4Darker50Percent,
        /// <summary>
        /// Accent 5 Darker 50%
        /// </summary>
        Accent5Darker50Percent,
        /// <summary>
        /// Accent 6 Darker 50%
        /// </summary>
        Accent6Darker50Percent,
        /// <summary>
        /// Accent 1 Darker 25%
        /// </summary>
        Accent1Darker25Percent,
        /// <summary>
        /// Accent 2 Darker 25%
        /// </summary>
        Accent2Darker25Percent,
        /// <summary>
        /// Accent 3 Darker 25%
        /// </summary>
        Accent3Darker25Percent,
        /// <summary>
        /// Accent 4 Darker 25%
        /// </summary>
        Accent4Darker25Percent,
        /// <summary>
        /// Accent 5 Darker 25%
        /// </summary>
        Accent5Darker25Percent,
        /// <summary>
        /// Accent 6 Darker 25%
        /// </summary>
        Accent6Darker25Percent,
        /// <summary>
        /// Accent 1
        /// </summary>
        Accent1,
        /// <summary>
        /// Accent 2
        /// </summary>
        Accent2,
        /// <summary>
        /// Accent 3
        /// </summary>
        Accent3,
        /// <summary>
        /// Accent 4
        /// </summary>
        Accent4,
        /// <summary>
        /// Accent 5
        /// </summary>
        Accent5,
        /// <summary>
        /// Accent 6
        /// </summary>
        Accent6,
        /// <summary>
        /// Accent 1 Lighter 40%
        /// </summary>
        Accent1Lighter40Percent,
        /// <summary>
        /// Accent 2 Lighter 40%
        /// </summary>
        Accent2Lighter40Percent,
        /// <summary>
        /// Accent 3 Lighter 40%
        /// </summary>
        Accent3Lighter40Percent,
        /// <summary>
        /// Accent 4 Lighter 40%
        /// </summary>
        Accent4Lighter40Percent,
        /// <summary>
        /// Accent 5 Lighter 40%
        /// </summary>
        Accent5Lighter40Percent,
        /// <summary>
        /// Accent 6 Lighter 40%
        /// </summary>
        Accent6Lighter40Percent,
        /// <summary>
        /// Dark #1
        /// </summary>
        Dark1,
        /// <summary>
        /// Dark #2
        /// </summary>
        Dark2,
        /// <summary>
        /// Dark #3
        /// </summary>
        Dark3,
        /// <summary>
        /// Dark #4
        /// </summary>
        Dark4,
        /// <summary>
        /// Dark #5
        /// </summary>
        Dark5,
        /// <summary>
        /// Dark #6
        /// </summary>
        Dark6,
        /// <summary>
        /// Colorful #1
        /// </summary>
        Colorful1,
        /// <summary>
        /// Colorful #2
        /// </summary>
        Colorful2,
        /// <summary>
        /// Colorful #3
        /// </summary>
        Colorful3,
        /// <summary>
        /// Colorful #4
        /// </summary>
        Colorful4,
        /// <summary>
        /// Colorful #5
        /// </summary>
        Colorful5,
        /// <summary>
        /// Colorful #6
        /// </summary>
        Colorful6
    }

    /// <summary>
    /// Encapsulates properties and methods for specifying sparklines.
    /// This simulates the DocumentFormat.OpenXml.Office2010.Excel.SparklineGroup class.
    /// </summary>
    public class SLSparklineGroup
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        // these are only used for setting location. They're not synchronised if the individual
        // sparkline changes cell references.
        internal string WorksheetName;
        internal int StartRowIndex;
        internal int StartColumnIndex;
        internal int EndRowIndex;
        internal int EndColumnIndex;

        /// <summary>
        /// The color for the main sparkline series.
        /// </summary>
        public SLColor SeriesColor { get; set; }

        /// <summary>
        /// The color for negative points.
        /// </summary>
        public SLColor NegativeColor { get; set; }

        /// <summary>
        /// The color for the axis.
        /// </summary>
        public SLColor AxisColor { get; set; }

        /// <summary>
        /// The color for markers.
        /// </summary>
        public SLColor MarkersColor { get; set; }

        /// <summary>
        /// The color for the first point.
        /// </summary>
        public SLColor FirstMarkerColor { get; set; }

        /// <summary>
        /// The color for the last point.
        /// </summary>
        public SLColor LastMarkerColor { get; set; }

        /// <summary>
        /// The color for the high point.
        /// </summary>
        public SLColor HighMarkerColor { get; set; }

        /// <summary>
        /// The color for the low point.
        /// </summary>
        public SLColor LowMarkerColor { get; set; }

        internal double ManualMax { get; set; }
        internal X14.SparklineAxisMinMaxValues MaxAxisType { get; set; }
        internal double ManualMin { get; set; }
        internal X14.SparklineAxisMinMaxValues MinAxisType { get; set; }

        internal decimal decLineWeight;
        /// <summary>
        /// Line weight for the sparkline group in points, ranging from 0 pt to 1584 pt (both inclusive).
        /// </summary>
        public decimal LineWeight
        {
            get { return decLineWeight; }
            set
            {
                decLineWeight = value;
                if (decLineWeight < 0) decLineWeight = 0;
                if (decLineWeight > 1584) decLineWeight = 1584;
            }
        }

        /// <summary>
        /// The type of sparkline. Use "Stacked" for "Win/Loss".
        /// </summary>
        public X14.SparklineTypeValues Type { get; set; }

        internal string DateWorksheetName;
        internal int DateStartRowIndex;
        internal int DateStartColumnIndex;
        internal int DateEndRowIndex;
        internal int DateEndColumnIndex;

        internal bool DateAxis { get; set; }

        /// <summary>
        /// The default is to show empty cells with a gap.
        /// </summary>
        public X14.DisplayBlanksAsValues ShowEmptyCellsAs { get; set; }

        /// <summary>
        /// Specifies if markers are shown.
        /// </summary>
        public bool ShowMarkers { get; set; }

        /// <summary>
        /// Specifies if the high point is shown.
        /// </summary>
        public bool ShowHighPoint { get; set; }

        /// <summary>
        /// Specifies if the low point is shown.
        /// </summary>
        public bool ShowLowPoint { get; set; }

        /// <summary>
        /// Specifies if the first point is shown.
        /// </summary>
        public bool ShowFirstPoint { get; set; }

        /// <summary>
        /// Specifies if the last point is shown.
        /// </summary>
        public bool ShowLastPoint { get; set; }

        /// <summary>
        /// Specifies if negative points are shown.
        /// </summary>
        public bool ShowNegativePoints { get; set; }

        /// <summary>
        /// Specifies is the horizontal axis is shown. This only appears if there's sparkline data crossing the zero point.
        /// </summary>
        public bool ShowAxis { get; set; }

        /// <summary>
        /// Specifies if hidden data is shown.
        /// </summary>
        public bool ShowHiddenData { get; set; }

        /// <summary>
        /// Plot data right-to-left.
        /// </summary>
        public bool RightToLeft { get; set; }

        // supposed to contain less than 2^31 sparklines. But I'm not gonna enforce this...
        // See documentation on CT_Sparklines for this.
        internal List<SLSparkline> Sparklines { get; set; }

        internal SLSparklineGroup(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
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
            this.WorksheetName = string.Empty;
            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;

            this.SeriesColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            this.NegativeColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            this.AxisColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            this.MarkersColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            this.FirstMarkerColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            this.LastMarkerColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            this.HighMarkerColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            this.LowMarkerColor = new SLColor(this.listThemeColors, this.listIndexedColors);

            this.ManualMax = 0;
            this.MaxAxisType = X14.SparklineAxisMinMaxValues.Individual;
            this.ManualMin = 0;
            this.MinAxisType = X14.SparklineAxisMinMaxValues.Individual;

            this.decLineWeight = 0.75m;

            this.Type = X14.SparklineTypeValues.Line;

            this.DateWorksheetName = string.Empty;
            this.DateStartRowIndex = 1;
            this.DateStartColumnIndex = 1;
            this.DateEndRowIndex = 1;
            this.DateEndColumnIndex = 1;
            this.DateAxis = false;

            this.ShowEmptyCellsAs = X14.DisplayBlanksAsValues.Gap;

            this.ShowMarkers = false;
            this.ShowHighPoint = false;
            this.ShowLowPoint = false;
            this.ShowFirstPoint = false;
            this.ShowLastPoint = false;
            this.ShowNegativePoints = false;
            this.ShowAxis = false;
            this.ShowHiddenData = false;
            this.RightToLeft = false;

            this.Sparklines = new List<SLSparkline>();
        }

        /// <summary>
        /// Set the location of the sparkline group given a cell reference. Use this if your data source is either 1 row of cells or 1 column of cells.
        /// </summary>
        /// <param name="CellReference">The cell reference such as "A1".</param>
        public void SetLocation(string CellReference)
        {
            // in case developers copy straight from the Excel dialog box...
            string sCellReference = CellReference.Replace("$", "");

            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(sCellReference, out iRowIndex, out iColumnIndex))
            {
                iRowIndex = -1;
                iColumnIndex = -1;
            }

            this.SetLocation(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex, true);
        }

        /// <summary>
        /// Set the location of the sparkline group given cell references of opposite cells in a cell range.
        /// Note that the cell range has to be a 1-dimensional vector, meaning it's either a single row or single column.
        /// Note also that the length of the vector must be equal to either the number of rows or number of columns in the data source range.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the location cell range, such as "A1". This is either the top-most or left-most cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the location cell range, such as "A1". This is either the bottom-most or right-most cell.</param>
        public void SetLocation(string StartCellReference, string EndCellReference)
        {
            // in case developers copy straight from the Excel dialog box...
            string sStartCellReference = StartCellReference.Replace("$", "");
            string sEndCellReference = EndCellReference.Replace("$", "");

            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(sStartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(sEndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            this.SetLocation(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, true);
        }

        /// <summary>
        /// Set the location of the sparkline group given cell references of opposite cells in a cell range.
        /// Note that the cell range has to be a 1-dimensional vector, meaning it's either a single row or single column.
        /// Note also that the length of the vector must be equal to either the number of rows or number of columns in the data source range.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the location cell range, such as "A1". This is either the top-most or left-most cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the location cell range, such as "A1". This is either the bottom-most or right-most cell.</param>
        /// <param name="RowsAsDataSeries">True if the data source has its series in rows. False if it's in columns. This only comes into play if the data source has the same number of rows as its columns.</param>
        public void SetLocation(string StartCellReference, string EndCellReference, bool RowsAsDataSeries)
        {
            // in case developers copy straight from the Excel dialog box...
            string sStartCellReference = StartCellReference.Replace("$", "");
            string sEndCellReference = EndCellReference.Replace("$", "");

            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(sStartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(sEndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            this.SetLocation(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, RowsAsDataSeries);
        }

        /// <summary>
        /// Set the location of the sparkline group given a row and column index. Use this if your data source is either 1 row of cells or 1 column of cells.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        public void SetLocation(int RowIndex, int ColumnIndex)
        {
            this.SetLocation(RowIndex, ColumnIndex, RowIndex, ColumnIndex, true);
        }

        /// <summary>
        /// Set the location of the sparkline group given row and column indices of opposite cells in a cell range.
        /// Note that the cell range has to be a 1-dimensional vector, meaning it's either a single row or single column.
        /// Note also that the length of the vector must be equal to either the number of rows or number of columns in the data source range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        public void SetLocation(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            this.SetLocation(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, true);
        }

        /// <summary>
        /// Set the location of the sparkline group given row and column indices of opposite cells in a cell range.
        /// Note that the cell range has to be a 1-dimensional vector, meaning it's either a single row or single column.
        /// Note also that the length of the vector must be equal to either the number of rows or number of columns in the data source range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <param name="RowsAsDataSeries">True if the data source has its series in rows. False if it's in columns. This only comes into play if the data source has the same number of rows as its columns.</param>
        public void SetLocation(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool RowsAsDataSeries)
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

            int iLocationRowDimension = iEndRowIndex - iStartRowIndex + 1;
            int iLocationColumnDimension = iEndColumnIndex - iStartColumnIndex + 1;

            // either the location row or column dimension must be 1. One of them has to be 1
            // for the location range to be valid. If there's an error, we'll just "shorten"
            // the smaller of the 2 dimensions. The Excel user interface has error dialog boxes
            // to warn the user. We don't have this luxury, so we'll make the best of things...
            if (iLocationRowDimension != 1 && iLocationColumnDimension != 1)
            {
                if (iLocationRowDimension < iLocationColumnDimension)
                {
                    iEndRowIndex = iStartRowIndex;
                    iLocationRowDimension = 1;
                }
                else
                {
                    iEndColumnIndex = iStartColumnIndex;
                    iLocationColumnDimension = 1;
                }
            }

            int iDataRowDimension = this.EndRowIndex - this.StartRowIndex + 1;
            int iDataColumnDimension = this.EndColumnIndex - this.StartColumnIndex + 1;

            bool bRowsAsDataSeries = true;
            int iMaxLocationDimension = 1;
            if (iLocationRowDimension >= iLocationColumnDimension)
            {
                iMaxLocationDimension = iLocationRowDimension;
                bRowsAsDataSeries = true;
            }
            else
            {
                iMaxLocationDimension = iLocationColumnDimension;
                bRowsAsDataSeries = false;
            }

            // If the data source has the same number of rows as its columns, the "default" is to use rows as data series,
            // unless otherwise stated. This is the "otherwise stated" part.
            if (iDataRowDimension == iDataColumnDimension)
            {
                bRowsAsDataSeries = RowsAsDataSeries;
            }

            // Furthermore, the "length" of the location range has to be either equal to
            // the data source range's row dimension or column dimension.
            // This is how Excel determines whether to use rows or columns as data series.

            int index;
            SLSparkline spk;
            if (iMaxLocationDimension == iDataRowDimension)
            {
                // sparkline data in row
                for (index = 0; index < iMaxLocationDimension; ++index)
                {
                    spk = new SLSparkline();
                    spk.WorksheetName = this.WorksheetName;
                    spk.StartRowIndex = index + this.StartRowIndex;
                    spk.EndRowIndex = spk.StartRowIndex;
                    spk.StartColumnIndex = this.StartColumnIndex;
                    spk.EndColumnIndex = this.EndColumnIndex;

                    if (bRowsAsDataSeries)
                    {
                        spk.LocationRowIndex = index + iStartRowIndex;
                        spk.LocationColumnIndex = iStartColumnIndex;
                    }
                    else
                    {
                        spk.LocationRowIndex = iStartRowIndex;
                        spk.LocationColumnIndex = index + iStartColumnIndex;
                    }

                    this.Sparklines.Add(spk);
                }
            }
            else if (iMaxLocationDimension == iDataColumnDimension)
            {
                // sparkline data in column
                for (index = 0; index < iMaxLocationDimension; ++index)
                {
                    spk = new SLSparkline();
                    spk.WorksheetName = this.WorksheetName;
                    spk.StartRowIndex = this.StartRowIndex;
                    spk.EndRowIndex = this.EndRowIndex;
                    spk.StartColumnIndex = index + this.StartColumnIndex;
                    spk.EndColumnIndex = spk.StartColumnIndex;

                    if (bRowsAsDataSeries)
                    {
                        spk.LocationRowIndex = index + iStartRowIndex;
                        spk.LocationColumnIndex = iStartColumnIndex;
                    }
                    else
                    {
                        spk.LocationRowIndex = iStartRowIndex;
                        spk.LocationColumnIndex = index + iStartColumnIndex;
                    }

                    this.Sparklines.Add(spk);
                }
            }
            else
            {
                // don't do anything? The location range is invalid to the point
                // where I don't know what to do. So just don't do anything...
            }
        }

        /// <summary>
        /// Set the horizontal axis as general axis type.
        /// </summary>
        public void SetGeneralAxis()
        {
            this.DateAxis = false;
        }

        /// <summary>
        /// Set the horizontal axis as date axis type, given a cell range containing the date values.
        /// Note that this means the cell range is a 1-dimensional vector, meaning it's a single row or single column.
        /// Note also that this probably means the length of the vector is the same as your location cell range.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the date cell range, such as "A1". This is either the top-most or left-most cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the date cell range, such as "A1". This is either the bottom-most or right-most cell.</param>
        public void SetDateAxis(string StartCellReference, string EndCellReference)
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

            this.SetDateAxis(this.WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Set the horizontal axis as date axis type, given a worksheet name and a cell range containing the date values.
        /// Note that this means the cell range is a 1-dimensional vector, meaning it's a single row or single column.
        /// Note also that this probably means the length of the vector is the same as your location cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartCellReference">The cell reference of the start cell of the date cell range, such as "A1". This is either the top-most or left-most cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the date cell range, such as "A1". This is either the bottom-most or right-most cell.</param>
        public void SetDateAxis(string WorksheetName, string StartCellReference, string EndCellReference)
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

            this.SetDateAxis(WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Set the horizontal axis as date axis type, given row and column indices of opposite cells in a cell range containing the date values.
        /// Note that this means the cell range is a 1-dimensional vector, meaning it's a single row or single column.
        /// Note also that this probably means the length of the vector is the same as your location cell range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        public void SetDateAxis(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            this.SetDateAxis(this.WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex);
        }

        /// <summary>
        /// Set the horizontal axis as date axis type, given a worksheet name, and row and column indices of opposite cells in a cell range containing the date values.
        /// Note that this means the cell range is a 1-dimensional vector, meaning it's a single row or single column.
        /// Note also that this probably means the length of the vector is the same as your location cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        public void SetDateAxis(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            this.DateAxis = true;

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

            this.DateWorksheetName = WorksheetName;
            this.DateStartRowIndex = iStartRowIndex;
            this.DateStartColumnIndex = iStartColumnIndex;
            this.DateEndRowIndex = iEndRowIndex;
            this.DateEndColumnIndex = iEndColumnIndex;
        }

        /// <summary>
        /// Set automatic minimum value for the vertical axis for the entire sparkline group.
        /// </summary>
        public void SetAutomaticMinimumValue()
        {
            this.MinAxisType = X14.SparklineAxisMinMaxValues.Individual;
            this.ManualMin = 0;
        }

        /// <summary>
        /// Set the same minimum value for the vertical axis for the entire sparkline group.
        /// </summary>
        public void SetSameMinimumValue()
        {
            this.MinAxisType = X14.SparklineAxisMinMaxValues.Group;
            this.ManualMin = 0;
        }

        /// <summary>
        /// Set a custom minimum value for the vertical axis for the entire sparkline group.
        /// </summary>
        /// <param name="MinValue">The custom minimum value.</param>
        public void SetCustomMinimumValue(double MinValue)
        {
            this.MinAxisType = X14.SparklineAxisMinMaxValues.Custom;
            this.ManualMin = MinValue;
        }

        /// <summary>
        /// Set automatic maximum value for the vertical axis for the entire sparkline group.
        /// </summary>
        public void SetAutomaticMaximumValue()
        {
            this.MaxAxisType = X14.SparklineAxisMinMaxValues.Individual;
            this.ManualMax = 0;
        }

        /// <summary>
        /// Set the same maximum value for the vertical axis for the entire sparkline group.
        /// </summary>
        public void SetSameMaximumValue()
        {
            this.MaxAxisType = X14.SparklineAxisMinMaxValues.Group;
            this.ManualMax = 0;
        }

        /// <summary>
        /// Set a custom maximum value for the vertical axis for the entire sparkline group.
        /// </summary>
        /// <param name="MaxValue">The custom maximum value.</param>
        public void SetCustomMaximumValue(double MaxValue)
        {
            this.MaxAxisType = X14.SparklineAxisMinMaxValues.Custom;
            this.ManualMax = MaxValue;
        }

        /// <summary>
        /// Set the sparkline style.
        /// </summary>
        /// <param name="Style">A built-in sparkline style.</param>
        public void SetSparklineStyle(SLSparklineStyle Style)
        {
            switch (Style)
            {
                case SLSparklineStyle.Accent1Darker50Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.499984740745262);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.499984740745262);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, 0.39997558519241921);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, 0.39997558519241921);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    break;
                case SLSparklineStyle.Accent2Darker50Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.499984740745262);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.499984740745262);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, 0.39997558519241921);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, 0.39997558519241921);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    break;
                case SLSparklineStyle.Accent3Darker50Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.499984740745262);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.499984740745262);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, 0.39997558519241921);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, 0.39997558519241921);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    break;
                case SLSparklineStyle.Accent4Darker50Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.499984740745262);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.499984740745262);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, 0.39997558519241921);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, 0.39997558519241921);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    break;
                case SLSparklineStyle.Accent5Darker50Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.499984740745262);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.499984740745262);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, 0.39997558519241921);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, 0.39997558519241921);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    break;
                case SLSparklineStyle.Accent6Darker50Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.499984740745262);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.499984740745262);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, 0.39997558519241921);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, 0.39997558519241921);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    break;
                case SLSparklineStyle.Accent1Darker25Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent2Darker25Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent3Darker25Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent4Darker25Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent5Darker25Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent6Darker25Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent1:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent2:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent3:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent4:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent5:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent6:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent1Lighter40Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, 0.39997558519241921);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, 0.79998168889431442);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.499984740745262);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent2Lighter40Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, 0.39997558519241921);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, 0.79998168889431442);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.499984740745262);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent3Lighter40Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, 0.39997558519241921);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, 0.79998168889431442);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.499984740745262);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent4Lighter40Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, 0.39997558519241921);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, 0.79998168889431442);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.499984740745262);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent5Lighter40Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, 0.39997558519241921);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, 0.79998168889431442);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.499984740745262);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent6Lighter40Percent:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, 0.39997558519241921);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, 0.79998168889431442);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.499984740745262);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Dark1:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.499984740745262);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    break;
                case SLSparklineStyle.Dark2:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.34998626667073579);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Dark3:
                    this.SeriesColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x32, 0x32, 0x32);
                    this.NegativeColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.FirstMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.LastMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.HighMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.LowMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    break;
                case SLSparklineStyle.Dark4:
                    this.SeriesColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.NegativeColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    this.FirstMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    this.LastMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    this.HighMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    this.LowMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    break;
                case SLSparklineStyle.Dark5:
                    this.SeriesColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x37, 0x60, 0x92);
                    this.NegativeColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.FirstMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.LastMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.HighMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    this.LowMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD0, 0, 0);
                    break;
                case SLSparklineStyle.Dark6:
                    this.SeriesColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    this.NegativeColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.FirstMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.LastMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.HighMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.LowMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    break;
                case SLSparklineStyle.Colorful1:
                    this.SeriesColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x5F, 0x5F, 0x5F);
                    this.NegativeColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xB6, 0x20);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD7, 0, 0x77);
                    this.FirstMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x56, 0x87, 0xC2);
                    this.LastMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x35, 0x9C, 0xEB);
                    this.HighMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x56, 0xBE, 0x79);
                    this.LowMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x50, 0x55);
                    break;
                case SLSparklineStyle.Colorful2:
                    this.SeriesColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x56, 0x87, 0xC2);
                    this.NegativeColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xB6, 0x20);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD7, 0, 0x77);
                    this.FirstMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x77, 0x77, 0x77);
                    this.LastMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x35, 0x9C, 0xEB);
                    this.HighMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x56, 0xBE, 0x79);
                    this.LowMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x50, 0x55);
                    break;
                case SLSparklineStyle.Colorful3:
                    this.SeriesColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xC6, 0xEF, 0xCE);
                    this.NegativeColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xC7, 0xCE);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x8C, 0xAD, 0xD6);
                    this.FirstMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xDC, 0x47);
                    this.LastMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEB, 0x9C);
                    this.HighMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x60, 0xD2, 0x76);
                    this.LowMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x53, 0x67);
                    break;
                case SLSparklineStyle.Colorful4:
                    this.SeriesColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0xB0, 0x50);
                    this.NegativeColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0, 0);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    this.FirstMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xC0, 0);
                    this.LastMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xC0, 0);
                    this.HighMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0xB0, 0x50);
                    this.LowMarkerColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0, 0);
                    break;
                case SLSparklineStyle.Colorful5:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Dark2Color);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    break;
                case SLSparklineStyle.Colorful6:
                    this.SeriesColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color);
                    this.NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0, 0, 0);
                    this.MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    this.FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    this.LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    this.HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    this.LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    break;
            }
        }

        internal void FromSparklineGroup(X14.SparklineGroup spkgrp)
        {
            this.SetAllNull();

            if (spkgrp.SeriesColor != null) this.SeriesColor.FromSeriesColor(spkgrp.SeriesColor);
            if (spkgrp.NegativeColor != null) this.NegativeColor.FromNegativeColor(spkgrp.NegativeColor);
            if (spkgrp.AxisColor != null) this.AxisColor.FromAxisColor(spkgrp.AxisColor);
            if (spkgrp.MarkersColor != null) this.MarkersColor.FromMarkersColor(spkgrp.MarkersColor);
            if (spkgrp.FirstMarkerColor != null) this.FirstMarkerColor.FromFirstMarkerColor(spkgrp.FirstMarkerColor);
            if (spkgrp.LastMarkerColor != null) this.LastMarkerColor.FromLastMarkerColor(spkgrp.LastMarkerColor);
            if (spkgrp.HighMarkerColor != null) this.HighMarkerColor.FromHighMarkerColor(spkgrp.HighMarkerColor);
            if (spkgrp.LowMarkerColor != null) this.LowMarkerColor.FromLowMarkerColor(spkgrp.LowMarkerColor);

            int index;
            string sRef = string.Empty;
            string sWorksheetName = string.Empty;
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;

            if (spkgrp.Formula != null)
            {
                sRef = spkgrp.Formula.Text;
                index = sRef.IndexOf("!");
                if (index >= 0)
                {
                    this.DateWorksheetName = sRef.Substring(0, index);
                    sRef = sRef.Substring(index + 1);
                }

                index = sRef.LastIndexOf(":");

                if (index >= 0)
                {
                    if (!SLTool.FormatCellReferenceRangeToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex, out iEndRowIndex, out iEndColumnIndex))
                    {
                        iStartRowIndex = -1;
                        iStartColumnIndex = -1;
                        iEndRowIndex = -1;
                        iEndColumnIndex = -1;
                    }

                    if (iStartRowIndex > 0 && iStartColumnIndex > 0 && iEndRowIndex > 0 && iEndColumnIndex > 0)
                    {
                        this.DateStartRowIndex = iStartRowIndex;
                        this.DateStartColumnIndex = iStartColumnIndex;
                        this.DateEndRowIndex = iEndRowIndex;
                        this.DateEndColumnIndex = iEndColumnIndex;
                        this.DateAxis = true;
                    }
                    else
                    {
                        this.DateStartRowIndex = 1;
                        this.DateStartColumnIndex = 1;
                        this.DateEndRowIndex = 1;
                        this.DateEndColumnIndex = 1;
                        this.DateAxis = false;
                    }
                }
                else
                {
                    if (!SLTool.FormatCellReferenceToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex))
                    {
                        iStartRowIndex = -1;
                        iStartColumnIndex = -1;
                    }

                    if (iStartRowIndex > 0 && iStartColumnIndex > 0)
                    {
                        this.DateStartRowIndex = iStartRowIndex;
                        this.DateStartColumnIndex = iStartColumnIndex;
                        this.DateEndRowIndex = iStartRowIndex;
                        this.DateEndColumnIndex = iStartColumnIndex;
                        this.DateAxis = true;
                    }
                    else
                    {
                        this.DateStartRowIndex = 1;
                        this.DateStartColumnIndex = 1;
                        this.DateEndRowIndex = 1;
                        this.DateEndColumnIndex = 1;
                        this.DateAxis = false;
                    }
                }
            }

            if (spkgrp.Sparklines != null)
            {
                X14.Sparkline spkline;
                SLSparkline spk;
                foreach (var child in spkgrp.Sparklines.ChildElements)
                {
                    if (child is X14.Sparkline)
                    {
                        spkline = (X14.Sparkline)child;
                        spk = new SLSparkline();
                        // the formula part contains the data source. Apparently, Excel is fine
                        // if it's empty. IF IT'S EMPTY THEN DELETE THE WHOLE SPARKLINE!
                        // Ok, I'm fine now... I'm gonna treat an empty Formula as "invalid".
                        if (spkline.Formula != null && spkline.ReferenceSequence != null)
                        {
                            sRef = spkline.Formula.Text;
                            index = sRef.IndexOf("!");
                            if (index >= 0)
                            {
                                spk.WorksheetName = sRef.Substring(0, index);
                                sRef = sRef.Substring(index + 1);
                            }

                            index = sRef.LastIndexOf(":");

                            if (index >= 0)
                            {
                                if (!SLTool.FormatCellReferenceRangeToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex, out iEndRowIndex, out iEndColumnIndex))
                                {
                                    iStartRowIndex = -1;
                                    iStartColumnIndex = -1;
                                    iEndRowIndex = -1;
                                    iEndColumnIndex = -1;
                                }

                                if (iStartRowIndex > 0 && iStartColumnIndex > 0 && iEndRowIndex > 0 && iEndColumnIndex > 0)
                                {
                                    spk.StartRowIndex = iStartRowIndex;
                                    spk.StartColumnIndex = iStartColumnIndex;
                                    spk.EndRowIndex = iEndRowIndex;
                                    spk.EndColumnIndex = iEndColumnIndex;
                                }
                                else
                                {
                                    spk.StartRowIndex = 1;
                                    spk.StartColumnIndex = 1;
                                    spk.EndRowIndex = 1;
                                    spk.EndColumnIndex = 1;
                                }
                            }
                            else
                            {
                                if (!SLTool.FormatCellReferenceToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex))
                                {
                                    iStartRowIndex = -1;
                                    iStartColumnIndex = -1;
                                }

                                if (iStartRowIndex > 0 && iStartColumnIndex > 0)
                                {
                                    spk.StartRowIndex = iStartRowIndex;
                                    spk.StartColumnIndex = iStartColumnIndex;
                                    spk.EndRowIndex = iStartRowIndex;
                                    spk.EndColumnIndex = iStartColumnIndex;
                                }
                                else
                                {
                                    spk.StartRowIndex = 1;
                                    spk.StartColumnIndex = 1;
                                    spk.EndRowIndex = 1;
                                    spk.EndColumnIndex = 1;
                                }
                            }

                            if (!SLTool.FormatCellReferenceToRowColumnIndex(spkline.ReferenceSequence.Text, out iStartRowIndex, out iStartColumnIndex))
                            {
                                iStartRowIndex = -1;
                                iStartColumnIndex = -1;
                            }

                            if (iStartRowIndex > 0 && iStartColumnIndex > 0)
                            {
                                spk.LocationRowIndex = iStartRowIndex;
                                spk.LocationColumnIndex = iStartColumnIndex;

                                // there are so many things that could possibly go wrong
                                // that we'll just assume that if the location part is correct,
                                // we'll just take it...
                                this.Sparklines.Add(spk.Clone());
                            }
                        }
                    }
                }
            }

            if (spkgrp.ManualMax != null) this.ManualMax = spkgrp.ManualMax.Value;
            if (spkgrp.ManualMin != null) this.ManualMin = spkgrp.ManualMin.Value;
            if (spkgrp.LineWeight != null) this.LineWeight = (decimal)spkgrp.LineWeight.Value;
            if (spkgrp.Type != null) this.Type = spkgrp.Type.Value;
            
            // we're gonna ignore dateAxis because if there's no formula, having it true is useless

            if (spkgrp.DisplayEmptyCellsAs != null) this.ShowEmptyCellsAs = spkgrp.DisplayEmptyCellsAs.Value;
            if (spkgrp.Markers != null) this.ShowMarkers = spkgrp.Markers.Value;
            if (spkgrp.High != null) this.ShowHighPoint = spkgrp.High.Value;
            if (spkgrp.Low != null) this.ShowLowPoint = spkgrp.Low.Value;
            if (spkgrp.First != null) this.ShowFirstPoint = spkgrp.First.Value;
            if (spkgrp.Last != null) this.ShowLastPoint = spkgrp.Last.Value;
            if (spkgrp.Negative != null) this.ShowNegativePoints = spkgrp.Negative.Value;
            if (spkgrp.DisplayXAxis != null) this.ShowAxis = spkgrp.DisplayXAxis.Value;
            if (spkgrp.DisplayHidden != null) this.ShowHiddenData = spkgrp.DisplayHidden.Value;
            if (spkgrp.MinAxisType != null) this.MinAxisType = spkgrp.MinAxisType.Value;
            if (spkgrp.MaxAxisType != null) this.MaxAxisType = spkgrp.MaxAxisType.Value;
            if (spkgrp.RightToLeft != null) this.RightToLeft = spkgrp.RightToLeft.Value;
        }

        internal X14.SparklineGroup ToSparklineGroup()
        {
            X14.SparklineGroup spkgrp = new X14.SparklineGroup();

            if (!this.SeriesColor.IsEmpty()) spkgrp.SeriesColor = this.SeriesColor.ToSeriesColor();
            if (!this.NegativeColor.IsEmpty()) spkgrp.NegativeColor = this.NegativeColor.ToNegativeColor();
            if (!this.AxisColor.IsEmpty()) spkgrp.AxisColor = this.AxisColor.ToAxisColor();
            if (!this.MarkersColor.IsEmpty()) spkgrp.MarkersColor = this.MarkersColor.ToMarkersColor();
            if (!this.FirstMarkerColor.IsEmpty()) spkgrp.FirstMarkerColor = this.FirstMarkerColor.ToFirstMarkerColor();
            if (!this.LastMarkerColor.IsEmpty()) spkgrp.LastMarkerColor = this.LastMarkerColor.ToLastMarkerColor();
            if (!this.HighMarkerColor.IsEmpty()) spkgrp.HighMarkerColor = this.HighMarkerColor.ToHighMarkerColor();
            if (!this.LowMarkerColor.IsEmpty()) spkgrp.LowMarkerColor = this.LowMarkerColor.ToLowMarkerColor();

            if (this.DateAxis)
            {
                if (this.DateStartRowIndex == this.DateEndRowIndex && this.DateStartColumnIndex == this.DateEndColumnIndex)
                {
                    spkgrp.Formula = new Excel.Formula();
                    spkgrp.Formula.Text = SLTool.ToCellReference(this.DateWorksheetName, this.DateStartRowIndex, this.DateStartColumnIndex);
                }
                else
                {
                    spkgrp.Formula = new Excel.Formula();
                    spkgrp.Formula.Text = SLTool.ToCellRange(this.DateWorksheetName, this.DateStartRowIndex, this.DateStartColumnIndex, this.DateEndRowIndex, this.DateEndColumnIndex);
                }

                spkgrp.DateAxis = true;
            }

            spkgrp.Sparklines = new X14.Sparklines();
            foreach (SLSparkline spk in this.Sparklines)
            {
                spkgrp.Sparklines.Append(spk.ToSparkline());
            }

            switch (this.MinAxisType)
            {
                case X14.SparklineAxisMinMaxValues.Individual:
                    // default, so don't have to do anything
                    break;
                case X14.SparklineAxisMinMaxValues.Group:
                    spkgrp.MinAxisType = X14.SparklineAxisMinMaxValues.Group;
                    break;
                case X14.SparklineAxisMinMaxValues.Custom:
                    spkgrp.MinAxisType = X14.SparklineAxisMinMaxValues.Custom;
                    spkgrp.ManualMin = this.ManualMin;
                    break;
            }

            switch (this.MaxAxisType)
            {
                case X14.SparklineAxisMinMaxValues.Individual:
                    // default, so don't have to do anything
                    break;
                case X14.SparklineAxisMinMaxValues.Group:
                    spkgrp.MaxAxisType = X14.SparklineAxisMinMaxValues.Group;
                    break;
                case X14.SparklineAxisMinMaxValues.Custom:
                    spkgrp.MaxAxisType = X14.SparklineAxisMinMaxValues.Custom;
                    spkgrp.ManualMax = this.ManualMax;
                    break;
            }

            if (this.decLineWeight != 0.75m) spkgrp.LineWeight = (double)this.decLineWeight;

            if (this.Type != X14.SparklineTypeValues.Line) spkgrp.Type = this.Type;

            if (this.ShowEmptyCellsAs != X14.DisplayBlanksAsValues.Zero) spkgrp.DisplayEmptyCellsAs = this.ShowEmptyCellsAs;

            if (this.ShowMarkers) spkgrp.Markers = true;
            if (this.ShowHighPoint) spkgrp.High = true;
            if (this.ShowLowPoint) spkgrp.Low = true;
            if (this.ShowFirstPoint) spkgrp.First = true;
            if (this.ShowLastPoint) spkgrp.Last = true;
            if (this.ShowNegativePoints) spkgrp.Negative = true;
            if (this.ShowAxis) spkgrp.DisplayXAxis = true;
            if (this.ShowHiddenData) spkgrp.DisplayHidden = true;
            if (this.RightToLeft) spkgrp.RightToLeft = true;

            return spkgrp;
        }

        internal SLSparklineGroup Clone()
        {
            SLSparklineGroup spkgrp = new SLSparklineGroup(this.listThemeColors, this.listIndexedColors);
            spkgrp.WorksheetName = this.WorksheetName;
            spkgrp.StartRowIndex = this.StartRowIndex;
            spkgrp.StartColumnIndex = this.StartColumnIndex;
            spkgrp.EndRowIndex = this.EndRowIndex;
            spkgrp.EndColumnIndex = this.EndColumnIndex;

            spkgrp.SeriesColor = this.SeriesColor.Clone();
            spkgrp.NegativeColor = this.NegativeColor.Clone();
            spkgrp.AxisColor = this.AxisColor.Clone();
            spkgrp.MarkersColor = this.MarkersColor.Clone();
            spkgrp.FirstMarkerColor = this.FirstMarkerColor.Clone();
            spkgrp.LastMarkerColor = this.LastMarkerColor.Clone();
            spkgrp.HighMarkerColor = this.HighMarkerColor.Clone();
            spkgrp.LowMarkerColor = this.LowMarkerColor.Clone();

            spkgrp.DateWorksheetName = this.DateWorksheetName;
            spkgrp.DateStartRowIndex = this.DateStartRowIndex;
            spkgrp.DateStartColumnIndex = this.DateStartColumnIndex;
            spkgrp.DateEndRowIndex = this.DateEndRowIndex;
            spkgrp.DateEndColumnIndex = this.DateEndColumnIndex;
            spkgrp.DateAxis = this.DateAxis;

            foreach (SLSparkline spk in this.Sparklines)
            {
                spkgrp.Sparklines.Add(spk.Clone());
            }

            spkgrp.ManualMax = this.ManualMax;
            spkgrp.MaxAxisType = this.MaxAxisType;
            spkgrp.ManualMin = this.ManualMin;
            spkgrp.MinAxisType = this.MinAxisType;

            spkgrp.decLineWeight = this.decLineWeight;

            spkgrp.Type = this.Type;

            spkgrp.ShowEmptyCellsAs = this.ShowEmptyCellsAs;

            spkgrp.ShowMarkers = this.ShowMarkers;
            spkgrp.ShowHighPoint = this.ShowHighPoint;
            spkgrp.ShowLowPoint = this.ShowLowPoint;
            spkgrp.ShowFirstPoint = this.ShowFirstPoint;
            spkgrp.ShowLastPoint = this.ShowLastPoint;
            spkgrp.ShowNegativePoints = this.ShowNegativePoints;
            spkgrp.ShowAxis = this.ShowAxis;
            spkgrp.ShowHiddenData = this.ShowHiddenData;
            spkgrp.RightToLeft = this.RightToLeft;            

            return spkgrp;
        }
    }
}
