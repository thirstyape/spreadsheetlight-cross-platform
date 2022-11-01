using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLSortCondition
    {
        internal bool? Descending { get; set; }

        internal bool HasSortBy;
        private SortByValues vSortBy;
        internal SortByValues SortBy
        {
            get { return vSortBy; }
            set
            {
                vSortBy = value;
                HasSortBy = vSortBy != SortByValues.Value ? true : false;
            }
        }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal string CustomList { get; set; }
        internal uint? FormatId { get; set; }

        internal bool HasIconSet;
        private IconSetValues vIconSet;
        internal IconSetValues IconSet
        {
            get { return vIconSet; }
            set
            {
                vIconSet = value;
                HasIconSet = vIconSet != IconSetValues.ThreeArrows ? true : false;
            }
        }

        internal uint? IconId { get; set; }

        internal SLSortCondition()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Descending = null;
            this.vSortBy = SortByValues.Value;
            this.HasSortBy = false;

            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;

            this.CustomList = null;
            this.FormatId = null;

            this.vIconSet = IconSetValues.ThreeArrows;
            this.HasIconSet = false;

            this.IconId = null;
        }

        internal void FromSortCondition(SortCondition sc)
        {
            this.SetAllNull();

            if (sc.Descending != null && sc.Descending.Value) this.Descending = sc.Descending.Value;
            if (sc.SortBy != null) this.SortBy = sc.SortBy.Value;

            int iStartRowIndex = 1;
            int iStartColumnIndex = 1;
            int iEndRowIndex = 1;
            int iEndColumnIndex = 1;
            string sRef = sc.Reference.Value;
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

            if (sc.CustomList != null) this.CustomList = sc.CustomList.Value;
            if (sc.FormatId != null) this.FormatId = sc.FormatId.Value;
            if (sc.IconSet != null) this.IconSet = sc.IconSet.Value;
            if (sc.IconId != null) this.IconId = sc.IconId.Value;
        }

        internal SortCondition ToSortCondition()
        {
            SortCondition sc = new SortCondition();
            if (this.Descending != null) sc.Descending = this.Descending.Value;
            if (HasSortBy) sc.SortBy = this.SortBy;

            if (this.StartRowIndex == this.EndRowIndex && this.StartColumnIndex == this.EndColumnIndex)
            {
                sc.Reference = SLTool.ToCellReference(this.StartRowIndex, this.StartColumnIndex);
            }
            else
            {
                sc.Reference = string.Format("{0}:{1}",
                    SLTool.ToCellReference(this.StartRowIndex, this.StartColumnIndex),
                    SLTool.ToCellReference(this.EndRowIndex, this.EndColumnIndex));
            }

            if (this.CustomList != null) sc.CustomList = this.CustomList;
            if (this.FormatId != null) sc.FormatId = this.FormatId;
            if (HasIconSet) sc.IconSet = this.IconSet;
            if (this.IconId != null) sc.IconId = this.IconId.Value;

            return sc;
        }

        internal SLSortCondition Clone()
        {
            SLSortCondition sc = new SLSortCondition();
            sc.Descending = this.Descending;
            sc.HasSortBy = this.HasSortBy;
            sc.vSortBy = this.vSortBy;
            sc.StartRowIndex = this.StartRowIndex;
            sc.StartColumnIndex = this.StartColumnIndex;
            sc.EndRowIndex = this.EndRowIndex;
            sc.EndColumnIndex = this.EndColumnIndex;
            sc.CustomList = this.CustomList;
            sc.FormatId = this.FormatId;
            sc.HasIconSet = this.HasIconSet;
            sc.vIconSet = this.vIconSet;
            sc.IconId = this.IconId;

            return sc;
        }
    }
}
