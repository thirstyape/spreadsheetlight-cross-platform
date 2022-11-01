using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLErrorItem
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

        internal SLErrorItem()
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

        internal void FromErrorItem(ErrorItem ei)
        {
            this.SetAllNull();

            if (ei.Val != null) this.Val = ei.Val.Value;
            if (ei.Unused != null) this.Unused = ei.Unused.Value;
            if (ei.Calculated != null) this.Calculated = ei.Calculated.Value;
            if (ei.Caption != null) this.Caption = ei.Caption.Value;
            if (ei.PropertyCount != null) this.PropertyCount = ei.PropertyCount.Value;
            if (ei.FormatIndex != null) this.FormatIndex = ei.FormatIndex.Value;
            if (ei.BackgroundColor != null) this.BackgroundColor = ei.BackgroundColor.Value;
            if (ei.ForegroundColor != null) this.ForegroundColor = ei.ForegroundColor.Value;
            if (ei.Italic != null) this.Italic = ei.Italic.Value;
            if (ei.Underline != null) this.Underline = ei.Underline.Value;
            if (ei.Strikethrough != null) this.Strikethrough = ei.Strikethrough.Value;
            if (ei.Bold != null) this.Bold = ei.Bold.Value;

            SLTuplesType tt;
            MemberPropertyIndex mpi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(ei))
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

        internal ErrorItem ToErrorItem()
        {
            ErrorItem ei = new ErrorItem();
            ei.Val = this.Val;
            if (this.Unused != null) ei.Unused = this.Unused.Value;
            if (this.Calculated != null) ei.Calculated = this.Calculated.Value;
            if (this.Caption != null && this.Caption.Length > 0) ei.Caption = this.Caption;
            if (this.PropertyCount != null) ei.PropertyCount = this.PropertyCount.Value;
            if (this.FormatIndex != null) ei.FormatIndex = this.FormatIndex.Value;
            if (this.BackgroundColor != null && this.BackgroundColor.Length > 0) ei.BackgroundColor = new HexBinaryValue(this.BackgroundColor);
            if (this.ForegroundColor != null && this.ForegroundColor.Length > 0) ei.ForegroundColor = new HexBinaryValue(this.ForegroundColor);
            if (this.Italic != false) ei.Italic = this.Italic;
            if (this.Underline != false) ei.Underline = this.Underline;
            if (this.Strikethrough != false) ei.Strikethrough = this.Strikethrough;
            if (this.Bold != false) ei.Bold = this.Bold;

            foreach (SLTuplesType tt in this.Tuples)
            {
                ei.Append(tt.ToTuples());
            }

            foreach (int i in this.MemberPropertyIndexes)
            {
                if (i != 0) ei.Append(new MemberPropertyIndex() { Val = i });
                else ei.Append(new MemberPropertyIndex());
            }

            return ei;
        }

        internal SLErrorItem Clone()
        {
            SLErrorItem ei = new SLErrorItem();
            ei.Val = this.Val;
            ei.Unused = this.Unused;
            ei.Calculated = this.Calculated;
            ei.Caption = this.Caption;
            ei.PropertyCount = this.PropertyCount;
            ei.FormatIndex = this.FormatIndex;
            ei.BackgroundColor = this.BackgroundColor;
            ei.ForegroundColor = this.ForegroundColor;
            ei.Italic = this.Italic;
            ei.Underline = this.Underline;
            ei.Strikethrough = this.Strikethrough;
            ei.Bold = this.Bold;

            ei.Tuples = new List<SLTuplesType>();
            foreach (SLTuplesType tt in this.Tuples)
            {
                ei.Tuples.Add(tt.Clone());
            }

            ei.MemberPropertyIndexes = new List<int>();
            foreach (int i in this.MemberPropertyIndexes)
            {
                ei.MemberPropertyIndexes.Add(i);
            }

            return ei;
        }
    }
}
