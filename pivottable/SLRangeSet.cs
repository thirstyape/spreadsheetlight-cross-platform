using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLRangeSet
    {
        internal uint? FieldItemIndexPage1 { get; set; }
        internal uint? FieldItemIndexPage2 { get; set; }
        internal uint? FieldItemIndexPage3 { get; set; }
        internal uint? FieldItemIndexPage4 { get; set; }
        internal string Reference { get; set; }
        internal string Name { get; set; }
        internal string Sheet { get; set; }
        internal string Id { get; set; }

        internal SLRangeSet()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.FieldItemIndexPage1 = null;
            this.FieldItemIndexPage2 = null;
            this.FieldItemIndexPage3 = null;
            this.FieldItemIndexPage4 = null;
            this.Reference = "";
            this.Name = "";
            this.Sheet = "";
            this.Id = "";
        }

        internal void FromRangeSet(RangeSet rs)
        {
            this.SetAllNull();

            if (rs.FieldItemIndexPage1 != null) this.FieldItemIndexPage1 = rs.FieldItemIndexPage1.Value;
            if (rs.FieldItemIndexPage2 != null) this.FieldItemIndexPage2 = rs.FieldItemIndexPage2.Value;
            if (rs.FieldItemIndexPage3 != null) this.FieldItemIndexPage3 = rs.FieldItemIndexPage3.Value;
            if (rs.FieldItemIndexPage4 != null) this.FieldItemIndexPage4 = rs.FieldItemIndexPage4.Value;
            if (rs.Reference != null) this.Reference = rs.Reference.Value;
            if (rs.Name != null) this.Name = rs.Name.Value;
            if (rs.Sheet != null) this.Sheet = rs.Sheet.Value;
            if (rs.Id != null) this.Id = rs.Id.Value;
        }

        internal RangeSet ToRangeSet()
        {
            RangeSet rs = new RangeSet();
            if (this.FieldItemIndexPage1 != null) rs.FieldItemIndexPage1 = this.FieldItemIndexPage1.Value;
            if (this.FieldItemIndexPage2 != null) rs.FieldItemIndexPage2 = this.FieldItemIndexPage2.Value;
            if (this.FieldItemIndexPage3 != null) rs.FieldItemIndexPage3 = this.FieldItemIndexPage3.Value;
            if (this.FieldItemIndexPage4 != null) rs.FieldItemIndexPage4 = this.FieldItemIndexPage4.Value;
            if (this.Reference != null && this.Reference.Length > 0) rs.Reference = this.Reference;
            if (this.Name != null && this.Name.Length > 0) rs.Name = this.Name;
            if (this.Sheet != null && this.Sheet.Length > 0) rs.Sheet = this.Sheet;
            if (this.Id != null && this.Id.Length > 0) rs.Id = this.Id;

            return rs;
        }

        internal SLRangeSet Clone()
        {
            SLRangeSet rs = new SLRangeSet();
            rs.FieldItemIndexPage1 = this.FieldItemIndexPage1;
            rs.FieldItemIndexPage2 = this.FieldItemIndexPage2;
            rs.FieldItemIndexPage3 = this.FieldItemIndexPage3;
            rs.FieldItemIndexPage4 = this.FieldItemIndexPage4;
            rs.Reference = this.Reference;
            rs.Name = this.Name;
            rs.Sheet = this.Sheet;
            rs.Id = this.Id;

            return rs;
        }
    }
}
