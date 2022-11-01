using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLSortState
    {
        internal List<SLSortCondition> SortConditions { get; set; }
        internal bool? ColumnSort { get; set; }
        internal bool? CaseSensitive { get; set; }

        internal bool HasSortMethod;
        private SortMethodValues vSortMethod;
        internal SortMethodValues SortMethod
        {
            get { return vSortMethod; }
            set
            {
                vSortMethod = value;
                HasSortMethod = vSortMethod != SortMethodValues.None ? true : false;
            }
        }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal SLSortState()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.SortConditions = new List<SLSortCondition>();
            this.ColumnSort = null;
            this.CaseSensitive = null;

            this.vSortMethod = SortMethodValues.None;
            this.HasSortMethod = false;

            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;
        }

        internal void FromSortState(SortState ss)
        {
            this.SetAllNull();

            if (ss.ColumnSort != null && ss.ColumnSort.Value) this.ColumnSort = ss.ColumnSort.Value;
            if (ss.CaseSensitive != null && ss.CaseSensitive.Value) this.CaseSensitive = ss.CaseSensitive.Value;
            if (ss.SortMethod != null) this.SortMethod = ss.SortMethod.Value;

            int iStartRowIndex = 1;
            int iStartColumnIndex = 1;
            int iEndRowIndex = 1;
            int iEndColumnIndex = 1;
            string sRef = ss.Reference.Value;
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

            if (ss.HasChildren)
            {
                SLSortCondition sc;
                using (OpenXmlReader oxr = OpenXmlReader.Create(ss))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(SortCondition))
                        {
                            sc = new SLSortCondition();
                            sc.FromSortCondition((SortCondition)oxr.LoadCurrentElement());
                            // limit of 64 from Open XML specs
                            if (this.SortConditions.Count < 64) this.SortConditions.Add(sc);
                        }
                    }
                }
            }
        }

        internal SortState ToSortState()
        {
            SortState ss = new SortState();
            if (this.ColumnSort != null && this.ColumnSort.Value) ss.ColumnSort = this.ColumnSort.Value;
            if (this.CaseSensitive != null && this.CaseSensitive.Value) ss.CaseSensitive = this.CaseSensitive.Value;
            if (HasSortMethod) ss.SortMethod = this.SortMethod;

            if (this.StartRowIndex == this.EndRowIndex && this.StartColumnIndex == this.EndColumnIndex)
            {
                ss.Reference = SLTool.ToCellReference(this.StartRowIndex, this.StartColumnIndex);
            }
            else
            {
                ss.Reference = string.Format("{0}:{1}",
                    SLTool.ToCellReference(this.StartRowIndex, this.StartColumnIndex),
                    SLTool.ToCellReference(this.EndRowIndex, this.EndColumnIndex));
            }

            if (this.SortConditions.Count > 0)
            {
                for (int i = 0; i < this.SortConditions.Count; ++i)
                {
                    ss.Append(this.SortConditions[i].ToSortCondition());
                }
            }

            return ss;
        }

        internal SLSortState Clone()
        {
            SLSortState ss = new SLSortState();
            ss.SortConditions = new List<SLSortCondition>();
            for (int i = 0; i < this.SortConditions.Count; ++i)
            {
                ss.SortConditions.Add(this.SortConditions[i].Clone());
            }

            ss.ColumnSort = this.ColumnSort;
            ss.CaseSensitive = this.CaseSensitive;

            ss.HasSortMethod = this.HasSortMethod;
            ss.vSortMethod = this.vSortMethod;

            ss.StartRowIndex = this.StartRowIndex;
            ss.StartColumnIndex = this.StartColumnIndex;
            ss.EndRowIndex = this.EndRowIndex;
            ss.EndColumnIndex = this.EndColumnIndex;

            return ss;
        }
    }
}
