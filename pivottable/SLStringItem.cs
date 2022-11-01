using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLStringItem
    {
        internal List<SLTuplesType> Tuples { get; set; }
        internal List<int> MemberPropertyIndexes { get; set; }

        internal string Val { get; set; }
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

        internal SLStringItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Tuples = new List<SLTuplesType>();
            this.MemberPropertyIndexes = new List<int>();

            this.Val = "";
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

        internal void FromStringItem(StringItem si)
        {
            this.SetAllNull();

            if (si.Val != null) this.Val = si.Val.Value;
            if (si.Unused != null) this.Unused = si.Unused.Value;
            if (si.Calculated != null) this.Calculated = si.Calculated.Value;
            if (si.Caption != null) this.Caption = si.Caption.Value;
            if (si.PropertyCount != null) this.PropertyCount = si.PropertyCount.Value;
            if (si.FormatIndex != null) this.FormatIndex = si.FormatIndex.Value;
            if (si.BackgroundColor != null) this.BackgroundColor = si.BackgroundColor.Value;
            if (si.ForegroundColor != null) this.ForegroundColor = si.ForegroundColor.Value;
            if (si.Italic != null) this.Italic = si.Italic.Value;
            if (si.Underline != null) this.Underline = si.Underline.Value;
            if (si.Strikethrough != null) this.Strikethrough = si.Strikethrough.Value;
            if (si.Bold != null) this.Bold = si.Bold.Value;

            SLTuplesType tt;
            MemberPropertyIndex mpi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(si))
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

        internal StringItem ToStringItem()
        {
            StringItem si = new StringItem();
            si.Val = this.Val;
            if (this.Unused != null) si.Unused = this.Unused.Value;
            if (this.Calculated != null) si.Calculated = this.Calculated.Value;
            if (this.Caption != null && this.Caption.Length > 0) si.Caption = this.Caption;
            if (this.PropertyCount != null) si.PropertyCount = this.PropertyCount.Value;
            if (this.FormatIndex != null) si.FormatIndex = this.FormatIndex.Value;
            if (this.BackgroundColor != null && this.BackgroundColor.Length > 0) si.BackgroundColor = new HexBinaryValue(this.BackgroundColor);
            if (this.ForegroundColor != null && this.ForegroundColor.Length > 0) si.ForegroundColor = new HexBinaryValue(this.ForegroundColor);
            if (this.Italic != false) si.Italic = this.Italic;
            if (this.Underline != false) si.Underline = this.Underline;
            if (this.Strikethrough != false) si.Strikethrough = this.Strikethrough;
            if (this.Bold != false) si.Bold = this.Bold;

            foreach (SLTuplesType tt in this.Tuples)
            {
                si.Append(tt.ToTuples());
            }

            foreach (int i in this.MemberPropertyIndexes)
            {
                if (i != 0) si.Append(new MemberPropertyIndex() { Val = i });
                else si.Append(new MemberPropertyIndex());
            }

            return si;
        }

        internal SLStringItem Clone()
        {
            SLStringItem si = new SLStringItem();
            si.Val = this.Val;
            si.Unused = this.Unused;
            si.Calculated = this.Calculated;
            si.Caption = this.Caption;
            si.PropertyCount = this.PropertyCount;
            si.FormatIndex = this.FormatIndex;
            si.BackgroundColor = this.BackgroundColor;
            si.ForegroundColor = this.ForegroundColor;
            si.Italic = this.Italic;
            si.Underline = this.Underline;
            si.Strikethrough = this.Strikethrough;
            si.Bold = this.Bold;

            si.Tuples = new List<SLTuplesType>();
            foreach (SLTuplesType tt in this.Tuples)
            {
                si.Tuples.Add(tt.Clone());
            }

            si.MemberPropertyIndexes = new List<int>();
            foreach (int i in this.MemberPropertyIndexes)
            {
                si.MemberPropertyIndexes.Add(i);
            }

            return si;
        }
    }
}
