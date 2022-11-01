using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLMissingItem
    {
        internal List<SLTuplesType> Tuples { get; set; }
        internal List<int> MemberPropertyIndexes { get; set; }

        internal bool? Unused { get; set; }
        internal bool? Calculated { get; set; }
        internal string Caption { get; set; }
        internal uint? PropertyCount { get; set; }
        internal uint? FormatIndex { get; set; }
        internal string BackgroundColor { get; set; }
        internal string ForegroundColor { get; set; }
        internal bool Italic { get; set; }
        internal bool Underline { get; set; }
        internal bool Strikethrough { get; set; }
        internal bool Bold { get; set; }

        internal SLMissingItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Tuples = new List<SLTuplesType>();
            this.MemberPropertyIndexes = new List<int>();

            this.Unused = null;
            this.Calculated = null;
            this.Caption = "";
            this.PropertyCount = null;
            this.FormatIndex = null;
            this.BackgroundColor = "";
            this.ForegroundColor = "";
            this.Italic = false;
            this.Underline = false;
            this.Strikethrough = false;
            this.Bold = false;
        }

        internal void FromMissingItem(MissingItem mi)
        {
            this.SetAllNull();

            if (mi.Unused != null) this.Unused = mi.Unused.Value;
            if (mi.Calculated != null) this.Calculated = mi.Calculated.Value;
            if (mi.Caption != null) this.Caption = mi.Caption.Value;
            if (mi.PropertyCount != null) this.PropertyCount = mi.PropertyCount.Value;
            if (mi.FormatIndex != null) this.FormatIndex = mi.FormatIndex.Value;
            if (mi.BackgroundColor != null) this.BackgroundColor = mi.BackgroundColor.Value;
            if (mi.ForegroundColor != null) this.ForegroundColor = mi.ForegroundColor.Value;
            if (mi.Italic != null) this.Italic = mi.Italic.Value;
            if (mi.Underline != null) this.Underline = mi.Underline.Value;
            if (mi.Strikethrough != null) this.Strikethrough = mi.Strikethrough.Value;
            if (mi.Bold != null) this.Bold = mi.Bold.Value;

            SLTuplesType tt;
            MemberPropertyIndex mpi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(mi))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Tuples))
                    {
                        tt = new SLTuplesType();
                        tt.FromTuples((Tuples)oxr.LoadCurrentElement());
                        this.Tuples.Add(tt);
                    }
                    else if (oxr.ElementType == typeof(MemberPropertyIndex))
                    {
                        // 0 is the default value.
                        mpi = (MemberPropertyIndex)oxr.LoadCurrentElement();
                        if (mpi.Val != null) this.MemberPropertyIndexes.Add(mpi.Val.Value);
                        else this.MemberPropertyIndexes.Add(0);
                    }
                }
            }
        }

        internal MissingItem ToMissingItem()
        {
            MissingItem mi = new MissingItem();
            if (this.Unused != null) mi.Unused = this.Unused.Value;
            if (this.Calculated != null) mi.Calculated = this.Calculated.Value;
            if (this.Caption != null && this.Caption.Length > 0) mi.Caption = this.Caption;
            if (this.PropertyCount != null) mi.PropertyCount = this.PropertyCount.Value;
            if (this.FormatIndex != null) mi.FormatIndex = this.FormatIndex.Value;
            if (this.BackgroundColor != null && this.BackgroundColor.Length > 0) mi.BackgroundColor = new HexBinaryValue(this.BackgroundColor);
            if (this.ForegroundColor != null && this.ForegroundColor.Length > 0) mi.ForegroundColor = new HexBinaryValue(this.ForegroundColor);
            if (this.Italic != false) mi.Italic = this.Italic;
            if (this.Underline != false) mi.Underline = this.Underline;
            if (this.Strikethrough != false) mi.Strikethrough = this.Strikethrough;
            if (this.Bold != false) mi.Bold = this.Bold;

            foreach (SLTuplesType tt in this.Tuples)
            {
                mi.Append(tt.ToTuples());
            }

            foreach (int i in this.MemberPropertyIndexes)
            {
                if (i != 0) mi.Append(new MemberPropertyIndex() { Val = i });
                else mi.Append(new MemberPropertyIndex());
            }

            return mi;
        }

        internal SLMissingItem Clone()
        {
            SLMissingItem mi = new SLMissingItem();
            mi.Unused = this.Unused;
            mi.Calculated = this.Calculated;
            mi.Caption = this.Caption;
            mi.PropertyCount = this.PropertyCount;
            mi.FormatIndex = this.FormatIndex;
            mi.BackgroundColor = this.BackgroundColor;
            mi.ForegroundColor = this.ForegroundColor;
            mi.Italic = this.Italic;
            mi.Underline = this.Underline;
            mi.Strikethrough = this.Strikethrough;
            mi.Bold = this.Bold;

            mi.Tuples = new List<SLTuplesType>();
            foreach (SLTuplesType tt in this.Tuples)
            {
                mi.Tuples.Add(tt.Clone());
            }

            mi.MemberPropertyIndexes = new List<int>();
            foreach (int i in this.MemberPropertyIndexes)
            {
                mi.MemberPropertyIndexes.Add(i);
            }

            return mi;
        }
    }
}
