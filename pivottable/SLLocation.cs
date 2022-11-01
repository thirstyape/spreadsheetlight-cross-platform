using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLLocation
    {
        internal SLCellPointRange Reference { get; set; }
        internal uint FirstHeaderRow { get; set; }
        internal uint FirstDataRow { get; set; }
        internal uint FirstDataColumn { get; set; }
        internal uint RowPageCount { get; set; }
        internal uint ColumnsPerPage { get; set; }

        internal SLLocation()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Reference = new SLCellPointRange(1, 1, 1, 1);
            this.FirstHeaderRow = 1;
            this.FirstDataRow = 1;
            this.FirstDataColumn = 1;
            this.RowPageCount = 0;
            this.ColumnsPerPage = 0;
        }

        internal void FromLocation(Location loc)
        {
            this.SetAllNull();

            if (loc.Reference != null) this.Reference = SLTool.TranslateReferenceToCellPointRange(loc.Reference.Value);
            if (loc.FirstHeaderRow != null) this.FirstHeaderRow = loc.FirstHeaderRow.Value;
            if (loc.FirstDataRow != null) this.FirstDataRow = loc.FirstDataRow.Value;
            if (loc.FirstDataColumn != null) this.FirstDataColumn = loc.FirstDataColumn.Value;
            if (loc.RowPageCount != null) this.RowPageCount = loc.RowPageCount.Value;
            if (loc.ColumnsPerPage != null) this.ColumnsPerPage = loc.ColumnsPerPage.Value;
        }

        internal Location ToLocation()
        {
            Location loc = new Location();
            if (this.Reference.StartRowIndex == this.Reference.EndRowIndex
                && this.Reference.StartColumnIndex == this.Reference.EndColumnIndex)
            {
                // this shouldn't happen because what's the point of a single cell as
                // the whole pivot table? Something's already very wrong by this time...
                loc.Reference = SLTool.ToCellReference(this.Reference.StartRowIndex, this.Reference.StartColumnIndex);
            }
            else
            {
                loc.Reference = SLTool.ToCellRange(this.Reference.StartRowIndex, this.Reference.StartColumnIndex, this.Reference.EndRowIndex, this.Reference.EndColumnIndex);
            }
            
            loc.FirstHeaderRow = this.FirstHeaderRow;
            loc.FirstDataRow = this.FirstDataRow;
            loc.FirstDataColumn = this.FirstDataColumn;
            if (this.RowPageCount != 0) loc.RowPageCount = this.RowPageCount;
            if (this.ColumnsPerPage != 0) loc.ColumnsPerPage = this.ColumnsPerPage;

            return loc;
        }

        internal SLLocation Clone()
        {
            SLLocation loc = new SLLocation();
            loc.Reference = new SLCellPointRange(this.Reference.StartRowIndex, this.Reference.StartColumnIndex, this.Reference.EndRowIndex, this.Reference.EndColumnIndex);
            loc.FirstHeaderRow = this.FirstHeaderRow;
            loc.FirstDataRow = this.FirstDataRow;
            loc.FirstDataColumn = this.FirstDataColumn;
            loc.RowPageCount = this.RowPageCount;
            loc.ColumnsPerPage = this.ColumnsPerPage;

            return loc;
        }
    }
}
