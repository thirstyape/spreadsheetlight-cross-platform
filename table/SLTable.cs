using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Totals row function types.
    /// </summary>
    public enum SLTotalsRowFunctionValues
    {
        /// <summary>
        /// Average
        /// </summary>
        Average = 0,
        /// <summary>
        /// Count non-empty cells
        /// </summary>
        Count,
        /// <summary>
        /// Count numbers
        /// </summary>
        CountNumbers,
        /// <summary>
        /// Maximum
        /// </summary>
        Maximum,
        /// <summary>
        /// Minimum
        /// </summary>
        Minimum,
        /// <summary>
        /// Standard deviation
        /// </summary>
        StandardDeviation,
        /// <summary>
        /// Sum
        /// </summary>
        Sum,
        /// <summary>
        /// Variance
        /// </summary>
        Variance
    }

    /// <summary>
    /// Encapsulates properties and methods for specifying tables. This simulates the DocumentFormat.OpenXml.Spreadsheet.Table class.
    /// </summary>
    public class SLTable
    {
        internal bool IsNewTable;
        internal string RelationshipID { get; set; }

        /// <summary>
        /// Indicates if the table has auto-filter.
        /// </summary>
        public bool HasAutoFilter { get; set; }

        internal SLAutoFilter AutoFilter { get; set; }
        internal bool HasSortState;
        internal SLSortState SortState { get; set; }

        internal List<SLTableColumn> TableColumns { get; set; }
        internal HashSet<string> TableNames { get; set; }

        internal bool HasTableStyleInfo;
        internal SLTableStyleInfo TableStyleInfo { get; set; }

        internal uint Id { get; set; }
        internal string Name { get; set; }

        internal string sDisplayName;
        /// <summary>
        /// There should be no spaces in the given value.
        /// Because display names of tables have to be unique across the entire spreadsheet,
        /// this can only be checked when the table is actually inserted into the worksheet.
        /// If the display name is duplicate, a new display name will be automatically assigned upon insertion.
        /// </summary>
        public string DisplayName
        {
            get { return sDisplayName; }
            set
            {
                sDisplayName = value;
                this.Name = sDisplayName;
            }
        }

        // The maximum length of this string should be 32,767 characters
        // We're not going to check this... TODO ?
        internal string Comment { get; set; }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal bool HasTableType;
        private TableValues vTableType;
        internal TableValues TableType
        {
            get { return vTableType; }
            set
            {
                vTableType = value;
                HasTableType = vTableType != TableValues.Worksheet ? true : false;
            }
        }

        internal uint HeaderRowCount { get; set; }
        internal bool? InsertRow { get; set; }
        internal bool? InsertRowShift { get; set; }

        internal uint TotalsRowCount { get; set; }
        internal bool? TotalsRowShown { get; set; }

        /// <summary>
        /// Indicates if the table has a totals row.
        /// </summary>
        public bool HasTotalRow
        {
            get { return this.TotalsRowCount > 0 ? true : false; }
            set
            {
                if (value)
                {
                    // When inserting the table, the full table collision checks will be done.
                    if (this.TotalsRowCount == 0)
                    {
                        int iTotalsRowIndex = this.EndRowIndex + 1;
                        if (iTotalsRowIndex <= SLConstants.RowLimit)
                        {
                            this.EndRowIndex += 1;
                            this.TotalsRowCount = 1;
                            this.TotalsRowShown = true;
                        }
                    }
                }
                else
                {
                    if (this.TotalsRowCount > 0)
                    {
                        this.EndRowIndex -= 1;
                        // keep it at least one row deep
                        if (this.HeaderRowCount > 0)
                        {
                            if (this.EndRowIndex <= this.StartRowIndex) this.EndRowIndex = this.StartRowIndex + 1;
                        }
                        else
                        {
                            if (this.EndRowIndex < this.StartRowIndex) this.EndRowIndex = this.StartRowIndex;
                        }

                        this.TotalsRowCount = 0;
                        // no need to set TotalsRowShown false because it's a historic flag.
                        // If the totals row is *ever* shown, set it to true.
                    }
                    // else totals row count is already zero
                }
            }
        }

        internal bool? Published { get; set; }
        internal uint? HeaderRowFormatId { get; set; }
        internal uint? DataFormatId { get; set; }
        internal uint? TotalsRowFormatId { get; set; }
        internal uint? HeaderRowBorderFormatId { get; set; }
        internal uint? BorderFormatId { get; set; }
        internal uint? TotalsRowBorderFormatId { get; set; }
        internal string HeaderRowCellStyle { get; set; }
        internal string DataCellStyle { get; set; }
        internal string TotalsRowCellStyle { get; set; }
        internal uint? ConnectionId { get; set; }

        /// <summary>
        /// Indicates if the table has banded rows.
        /// </summary>
        public bool HasBandedRows
        {
            get
            {
                // we'll default to true
                if (this.HasTableStyleInfo)
                {
                    return this.TableStyleInfo.ShowRowStripes != null ? this.TableStyleInfo.ShowRowStripes.Value : true;
                }
                else
                {
                    return true;
                }
            }
            set
            {
                this.TableStyleInfo.ShowRowStripes = value;
                this.HasTableStyleInfo = true;
            }
        }

        /// <summary>
        /// Indicates if the table has banded columns.
        /// </summary>
        public bool HasBandedColumns
        {
            get
            {
                // we'll default to false
                if (this.HasTableStyleInfo)
                {
                    return this.TableStyleInfo.ShowColumnStripes != null ? this.TableStyleInfo.ShowColumnStripes.Value : false;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                this.TableStyleInfo.ShowColumnStripes = value;
                this.HasTableStyleInfo = true;
            }
        }

        /// <summary>
        /// Indicates if the table has special formatting for the first column.
        /// </summary>
        public bool HasFirstColumnStyled
        {
            get
            {
                // we'll default to false
                if (this.HasTableStyleInfo)
                {
                    return this.TableStyleInfo.ShowFirstColumn != null ? this.TableStyleInfo.ShowFirstColumn.Value : false;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                this.TableStyleInfo.ShowFirstColumn = value;
                this.HasTableStyleInfo = true;
            }
        }

        /// <summary>
        /// Indicates if the table has special formatting for the last column.
        /// </summary>
        public bool HasLastColumnStyled
        {
            get
            {
                // we'll default to false
                if (this.HasTableStyleInfo)
                {
                    return this.TableStyleInfo.ShowLastColumn != null ? this.TableStyleInfo.ShowLastColumn.Value : false;
                }
                else
                {
                    return false;
                }
            }
            set
            {
                this.TableStyleInfo.ShowLastColumn = value;
                this.HasTableStyleInfo = true;
            }
        }

        internal SLTable()
        {
            this.SetAllNull();
        }

        internal void SetAllNull()
        {
            this.IsNewTable = true;
            this.RelationshipID = string.Empty;

            this.AutoFilter = new SLAutoFilter();
            this.HasAutoFilter = false;
            this.SortState = new SLSortState();
            this.HasSortState = false;
            this.TableColumns = new List<SLTableColumn>();
            this.TableNames = new HashSet<string>();
            this.TableStyleInfo = new SLTableStyleInfo();
            this.HasTableStyleInfo = false;

            this.Id = 0;
            this.Name = null;
            this.sDisplayName = string.Empty;
            this.Comment = null;
            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;
            this.TableType = TableValues.Worksheet;
            this.HasTableType = false;
            this.HeaderRowCount = 1;
            this.InsertRow = null;
            this.InsertRowShift = null;
            this.TotalsRowCount = 0;
            this.TotalsRowShown = null;
            this.Published = null;
            this.HeaderRowFormatId = null;
            this.DataFormatId = null;
            this.TotalsRowFormatId = null;
            this.HeaderRowBorderFormatId = null;
            this.BorderFormatId = null;
            this.TotalsRowBorderFormatId = null;
            this.HeaderRowCellStyle = null;
            this.DataCellStyle = null;
            this.TotalsRowCellStyle = null;
            this.ConnectionId = null;
        }

        /// <summary>
        /// Set the table style with a built-in style.
        /// </summary>
        /// <param name="TableStyle">A built-in table style.</param>
        public void SetTableStyle(SLTableStyleTypeValues TableStyle)
        {
            this.TableStyleInfo.SetTableStyle(TableStyle);
            this.HasTableStyleInfo = true;
        }

        /// <summary>
        /// Remove the label text or function in the totals row.
        /// </summary>
        /// <param name="TableColumnIndex">The table column index. For example, 1 for the 1st table column, 2 for the 2nd table column and so on.</param>
        public void RemoveTotalRowLabelFunction(int TableColumnIndex)
        {
            --TableColumnIndex;
            if (TableColumnIndex < 0 || TableColumnIndex >= this.TableColumns.Count) return;

            this.TableColumns[TableColumnIndex].TotalsRowLabel = null;
            this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.None;
            this.TableColumns[TableColumnIndex].HasTotalsRowFunction = false;
        }

        /// <summary>
        /// Set the label text in the totals row. Be sure to set <see cref="HasTotalRow"/> true first.
        /// </summary>
        /// <param name="TableColumnIndex">The table column index. For example, 1 for the 1st table column, 2 for the 2nd table column and so on.</param>
        /// <param name="Label">The label text.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetTotalRowLabel(int TableColumnIndex, string Label)
        {
            if (this.TotalsRowCount > 0)
            {
                --TableColumnIndex;
                if (TableColumnIndex < 0 || TableColumnIndex >= this.TableColumns.Count) return false;

                this.TableColumns[TableColumnIndex].TotalsRowLabel = Label;
                this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.None;
                this.TableColumns[TableColumnIndex].HasTotalsRowFunction = false;

                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Set the function in the totals row. Be sure to set <see cref="HasTotalRow"/> true first.
        /// </summary>
        /// <param name="TableColumnIndex">The table column index. For example, 1 for the 1st table column, 2 for the 2nd table column and so on.</param>
        /// <param name="TotalsRowFunction">The function type.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetTotalRowFunction(int TableColumnIndex, SLTotalsRowFunctionValues TotalsRowFunction)
        {
            if (this.TotalsRowCount > 0)
            {
                --TableColumnIndex;
                if (TableColumnIndex < 0 || TableColumnIndex >= this.TableColumns.Count) return false;

                this.TableColumns[TableColumnIndex].TotalsRowLabel = null;

                int iStartRowIndex = -1;
                int iEndRowIndex = -1;
                if (this.HeaderRowCount > 0) iStartRowIndex = this.StartRowIndex + 1;
                else iStartRowIndex = this.StartRowIndex;
                // not inclusive of the last totals row
                iEndRowIndex = this.EndRowIndex - 1;

                int iColumnIndex = this.StartColumnIndex + TableColumnIndex;

                switch (TotalsRowFunction)
                {
                    case SLTotalsRowFunctionValues.Average:
                        this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Average;
                        break;
                    case SLTotalsRowFunctionValues.Count:
                        this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Count;
                        break;
                    case SLTotalsRowFunctionValues.CountNumbers:
                        this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.CountNumbers;
                        break;
                    case SLTotalsRowFunctionValues.Maximum:
                        this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Maximum;
                        break;
                    case SLTotalsRowFunctionValues.Minimum:
                        this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Minimum;
                        break;
                    case SLTotalsRowFunctionValues.StandardDeviation:
                        this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.StandardDeviation;
                        break;
                    case SLTotalsRowFunctionValues.Sum:
                        this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Sum;
                        break;
                    case SLTotalsRowFunctionValues.Variance:
                        this.TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Variance;
                        break;
                }
                this.TableColumns[TableColumnIndex].HasTotalsRowFunction = true;

                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// To sort data within the table. Note that the sorting is only done when the table is inserted into the worksheet.
        /// </summary>
        /// <param name="TableColumnIndex">The table column index. For example, 1 for the 1st table column, 2 for the 2nd table column and so on.</param>
        /// <param name="SortAscending">True to sort in ascending order. False to sort in descending order.</param>
        public void Sort(int TableColumnIndex, bool SortAscending)
        {
            --TableColumnIndex;
            if (TableColumnIndex < 0 || TableColumnIndex >= this.TableColumns.Count) return;

            int iStartRowIndex = -1;
            int iEndRowIndex = -1;
            if (this.HeaderRowCount > 0) iStartRowIndex = this.StartRowIndex + 1;
            else iStartRowIndex = this.StartRowIndex;
            // not inclusive of the last totals row
            if (this.TotalsRowCount > 0) iEndRowIndex = this.EndRowIndex - 1;
            else iEndRowIndex = this.EndRowIndex;

            this.SortState = new SLSortState();
            this.SortState.StartRowIndex = iStartRowIndex;
            this.SortState.EndRowIndex = iEndRowIndex;
            this.SortState.StartColumnIndex = this.StartColumnIndex;
            this.SortState.EndColumnIndex = this.EndColumnIndex;

            SLSortCondition sc = new SLSortCondition();
            sc.StartRowIndex = iStartRowIndex;
            sc.StartColumnIndex = this.StartColumnIndex + TableColumnIndex;
            sc.EndRowIndex = iEndRowIndex;
            sc.EndColumnIndex = sc.StartColumnIndex;
            if (!SortAscending) sc.Descending = true;
            this.SortState.SortConditions.Add(sc);

            this.HasSortState = true;
        }

        internal void FromTable(Table t)
        {
            this.SetAllNull();

            if (t.AutoFilter != null)
            {
                this.AutoFilter.FromAutoFilter(t.AutoFilter);
                this.HasAutoFilter = true;
            }
            if (t.SortState != null)
            {
                this.SortState.FromSortState(t.SortState);
                this.HasSortState = true;
            }
            using (OpenXmlReader oxr = OpenXmlReader.Create(t.TableColumns))
            {
                SLTableColumn tc;
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(TableColumn))
                    {
                        tc = new SLTableColumn();
                        tc.FromTableColumn((TableColumn)oxr.LoadCurrentElement());
                        this.TableColumns.Add(tc);
                    }
                }
            }
            if (t.TableStyleInfo != null)
            {
                this.TableStyleInfo.FromTableStyleInfo(t.TableStyleInfo);
                this.HasTableStyleInfo = true;
            }

            this.Id = t.Id.Value;
            if (t.Name != null) this.Name = t.Name.Value;
            this.sDisplayName = t.DisplayName.Value;
            if (t.Comment != null) this.Comment = t.Comment.Value;

            int iStartRowIndex = 1;
            int iStartColumnIndex = 1;
            int iEndRowIndex = 1;
            int iEndColumnIndex = 1;
            string sRef = t.Reference.Value;
            if (sRef.IndexOf(":") > 0)
            {
                if (SLTool.FormatCellReferenceRangeToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex, out iEndRowIndex, out iEndColumnIndex))
                {
                    this.StartRowIndex = iStartRowIndex;
                    this.StartColumnIndex = iStartColumnIndex;
                    this.EndRowIndex = iEndRowIndex;
                    this.EndColumnIndex = iEndColumnIndex;
                }
            }
            else
            {
                if (SLTool.FormatCellReferenceToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex))
                {
                    this.StartRowIndex = iStartRowIndex;
                    this.StartColumnIndex = iStartColumnIndex;
                    this.EndRowIndex = iStartRowIndex;
                    this.EndColumnIndex = iStartColumnIndex;
                }
            }

            if (t.TableType != null) this.TableType = t.TableType.Value;
            if (t.HeaderRowCount != null && t.HeaderRowCount.Value != 1) this.HeaderRowCount = t.HeaderRowCount.Value;
            if (t.InsertRow != null && t.InsertRow.Value) this.InsertRow = t.InsertRow.Value;
            if (t.InsertRowShift != null && t.InsertRowShift.Value) this.InsertRowShift = t.InsertRowShift.Value;
            if (t.TotalsRowCount != null && t.TotalsRowCount.Value != 0) this.TotalsRowCount = t.TotalsRowCount.Value;
            if (t.TotalsRowShown != null && !t.TotalsRowShown.Value) this.TotalsRowShown = t.TotalsRowShown.Value;
            if (t.Published != null && t.Published.Value) this.Published = t.Published.Value;
            if (t.HeaderRowFormatId != null) this.HeaderRowFormatId = t.HeaderRowFormatId.Value;
            if (t.DataFormatId != null) this.DataFormatId = t.DataFormatId.Value;
            if (t.TotalsRowFormatId != null) this.TotalsRowFormatId = t.TotalsRowFormatId.Value;
            if (t.HeaderRowBorderFormatId != null) this.HeaderRowBorderFormatId = t.HeaderRowBorderFormatId.Value;
            if (t.BorderFormatId != null) this.BorderFormatId = t.BorderFormatId.Value;
            if (t.TotalsRowBorderFormatId != null) this.TotalsRowBorderFormatId = t.TotalsRowBorderFormatId.Value;
            if (t.HeaderRowCellStyle != null) this.HeaderRowCellStyle = t.HeaderRowCellStyle.Value;
            if (t.DataCellStyle != null) this.DataCellStyle = t.DataCellStyle.Value;
            if (t.TotalsRowCellStyle != null) this.TotalsRowCellStyle = t.TotalsRowCellStyle.Value;
            if (t.ConnectionId != null) this.ConnectionId = t.ConnectionId.Value;
        }

        internal Table ToTable()
        {
            Table t = new Table();
            if (HasAutoFilter) t.AutoFilter = this.AutoFilter.ToAutoFilter();
            if (HasSortState) t.SortState = this.SortState.ToSortState();

            t.TableColumns = new DocumentFormat.OpenXml.Spreadsheet.TableColumns() { Count = (uint)this.TableColumns.Count };
            for (int i = 0; i < this.TableColumns.Count; ++i)
            {
                t.TableColumns.Append(this.TableColumns[i].ToTableColumn());
            }

            if (HasTableStyleInfo) t.TableStyleInfo = this.TableStyleInfo.ToTableStyleInfo();

            t.Id = this.Id;
            if (this.Name != null) t.Name = this.Name;
            t.DisplayName = this.DisplayName;
            if (this.Comment != null) t.Comment = this.Comment;

            if (this.StartRowIndex == this.EndRowIndex && this.StartColumnIndex == this.EndColumnIndex)
            {
                t.Reference = SLTool.ToCellReference(this.StartRowIndex, this.StartColumnIndex);
            }
            else
            {
                t.Reference = string.Format("{0}:{1}",
                    SLTool.ToCellReference(this.StartRowIndex, this.StartColumnIndex),
                    SLTool.ToCellReference(this.EndRowIndex, this.EndColumnIndex));
            }

            if (HasTableType) t.TableType = this.TableType;
            if (this.HeaderRowCount != 1) t.HeaderRowCount = this.HeaderRowCount;
            if (this.InsertRow != null && this.InsertRow.Value) t.InsertRow = this.InsertRow.Value;
            if (this.InsertRowShift != null && this.InsertRowShift.Value) t.InsertRowShift = this.InsertRowShift.Value;
            if (this.TotalsRowCount != 0) t.TotalsRowCount = this.TotalsRowCount;
            if (this.TotalsRowShown != null && !this.TotalsRowShown.Value) t.TotalsRowShown = this.TotalsRowShown.Value;
            if (this.Published != null && this.Published.Value) t.Published = this.Published.Value;
            if (this.HeaderRowFormatId != null) t.HeaderRowFormatId = this.HeaderRowFormatId.Value;
            if (this.DataFormatId != null) t.DataFormatId = this.DataFormatId.Value;
            if (this.TotalsRowFormatId != null) t.TotalsRowFormatId = this.TotalsRowFormatId.Value;
            if (this.HeaderRowBorderFormatId != null) t.HeaderRowBorderFormatId = this.HeaderRowBorderFormatId.Value;
            if (this.BorderFormatId != null) t.BorderFormatId = this.BorderFormatId.Value;
            if (this.TotalsRowBorderFormatId != null) t.TotalsRowBorderFormatId = this.TotalsRowBorderFormatId.Value;
            if (this.HeaderRowCellStyle != null) t.HeaderRowCellStyle = this.HeaderRowCellStyle;
            if (this.DataCellStyle != null) t.DataCellStyle = this.DataCellStyle;
            if (this.TotalsRowCellStyle != null) t.TotalsRowCellStyle = this.TotalsRowCellStyle;
            if (this.ConnectionId != null) t.ConnectionId = this.ConnectionId.Value;

            return t;
        }

        internal SLTable Clone()
        {
            SLTable t = new SLTable();
            t.IsNewTable = this.IsNewTable;
            t.RelationshipID = this.RelationshipID;
            t.HasAutoFilter = this.HasAutoFilter;
            t.AutoFilter = this.AutoFilter.Clone();
            t.HasSortState = this.HasSortState;
            t.SortState = this.SortState.Clone();

            t.TableColumns = new List<SLTableColumn>();
            for (int i = 0; i < this.TableColumns.Count; ++i)
            {
                t.TableColumns.Add(this.TableColumns[i].Clone());
            }

            t.TableNames = new HashSet<string>();
            foreach (string s in this.TableNames)
            {
                t.TableNames.Add(s);
            }

            t.HasTableStyleInfo = this.HasTableStyleInfo;
            t.TableStyleInfo = this.TableStyleInfo.Clone();

            t.Id = this.Id;
            t.Name = this.Name;
            t.sDisplayName = this.sDisplayName;
            t.Comment = this.Comment;
            t.StartRowIndex = this.StartRowIndex;
            t.StartColumnIndex = this.StartColumnIndex;
            t.EndRowIndex = this.EndRowIndex;
            t.EndColumnIndex = this.EndColumnIndex;

            t.HasTableType = this.HasTableType;
            t.vTableType = this.vTableType;

            t.HeaderRowCount = this.HeaderRowCount;
            t.InsertRow = this.InsertRow;
            t.InsertRowShift = this.InsertRowShift;
            t.TotalsRowCount = this.TotalsRowCount;
            t.TotalsRowShown = this.TotalsRowShown;

            t.Published = this.Published;
            t.HeaderRowFormatId = this.HeaderRowFormatId;
            t.DataFormatId = this.DataFormatId;
            t.TotalsRowFormatId = this.TotalsRowFormatId;
            t.HeaderRowBorderFormatId = this.HeaderRowBorderFormatId;
            t.BorderFormatId = this.BorderFormatId;
            t.TotalsRowBorderFormatId = this.TotalsRowBorderFormatId;
            t.HeaderRowCellStyle = this.HeaderRowCellStyle;
            t.DataCellStyle = this.DataCellStyle;
            t.TotalsRowCellStyle = this.TotalsRowCellStyle;
            t.ConnectionId = this.ConnectionId;

            return t;
        }
    }
}
