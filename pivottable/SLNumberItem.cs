using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLNumberItem
    {
        internal List<SLTuplesType> Tuples { get; set; }
        internal List<int> MemberPropertyIndexes { get; set; }

        internal double Val { get; set; }
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

        internal SLNumberItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Tuples = new List<SLTuplesType>();
            this.MemberPropertyIndexes = new List<int>();

            this.Val = 0;
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

        internal void FromNumberItem(NumberItem ni)
        {
            this.SetAllNull();

            if (ni.Val != null) this.Val = ni.Val.Value;
            if (ni.Unused != null) this.Unused = ni.Unused.Value;
            if (ni.Calculated != null) this.Calculated = ni.Calculated.Value;
            if (ni.Caption != null) this.Caption = ni.Caption.Value;
            if (ni.PropertyCount != null) this.PropertyCount = ni.PropertyCount.Value;
            if (ni.FormatIndex != null) this.FormatIndex = ni.FormatIndex.Value;
            if (ni.BackgroundColor != null) this.BackgroundColor = ni.BackgroundColor.Value;
            if (ni.ForegroundColor != null) this.ForegroundColor = ni.ForegroundColor.Value;
            if (ni.Italic != null) this.Italic = ni.Italic.Value;
            if (ni.Underline != null) this.Underline = ni.Underline.Value;
            if (ni.Strikethrough != null) this.Strikethrough = ni.Strikethrough.Value;
            if (ni.Bold != null) this.Bold = ni.Bold.Value;

            SLTuplesType tt;
            MemberPropertyIndex mpi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(ni))
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

        internal NumberItem ToNumberItem()
        {
            NumberItem ni = new NumberItem();
            ni.Val = this.Val;
            if (this.Unused != null) ni.Unused = this.Unused.Value;
            if (this.Calculated != null) ni.Calculated = this.Calculated.Value;
            if (this.Caption != null && this.Caption.Length > 0) ni.Caption = this.Caption;
            if (this.PropertyCount != null) ni.PropertyCount = this.PropertyCount.Value;
            if (this.FormatIndex != null) ni.FormatIndex = this.FormatIndex.Value;
            if (this.BackgroundColor != null && this.BackgroundColor.Length > 0) ni.BackgroundColor = new HexBinaryValue(this.BackgroundColor);
            if (this.ForegroundColor != null && this.ForegroundColor.Length > 0) ni.ForegroundColor = new HexBinaryValue(this.ForegroundColor);
            if (this.Italic != false) ni.Italic = this.Italic;
            if (this.Underline != false) ni.Underline = this.Underline;
            if (this.Strikethrough != false) ni.Strikethrough = this.Strikethrough;
            if (this.Bold != false) ni.Bold = this.Bold;

            foreach (SLTuplesType tt in this.Tuples)
            {
                ni.Append(tt.ToTuples());
            }

            foreach (int i in this.MemberPropertyIndexes)
            {
                if (i != 0) ni.Append(new MemberPropertyIndex() { Val = i });
                else ni.Append(new MemberPropertyIndex());
            }

            return ni;
        }

        internal SLNumberItem Clone()
        {
            SLNumberItem ni = new SLNumberItem();
            ni.Val = this.Val;
            ni.Unused = this.Unused;
            ni.Calculated = this.Calculated;
            ni.Caption = this.Caption;
            ni.PropertyCount = this.PropertyCount;
            ni.FormatIndex = this.FormatIndex;
            ni.BackgroundColor = this.BackgroundColor;
            ni.ForegroundColor = this.ForegroundColor;
            ni.Italic = this.Italic;
            ni.Underline = this.Underline;
            ni.Strikethrough = this.Strikethrough;
            ni.Bold = this.Bold;

            ni.Tuples = new List<SLTuplesType>();
            foreach (SLTuplesType tt in this.Tuples)
            {
                ni.Tuples.Add(tt.Clone());
            }

            ni.MemberPropertyIndexes = new List<int>();
            foreach (int i in this.MemberPropertyIndexes)
            {
                ni.MemberPropertyIndexes.Add(i);
            }

            return ni;
        }
    }
}
