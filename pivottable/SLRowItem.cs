using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLRowItem
    {
        internal List<int> MemberPropertyIndexes { get; set; }

        internal ItemValues ItemType { get; set; }
        internal uint RepeatedItemCount { get; set; }
        internal uint Index { get; set; }

        internal SLRowItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.ItemType = ItemValues.Data;
            this.RepeatedItemCount = 0;
            this.Index = 0;
        }

        internal void FromRowItem(RowItem ri)
        {
            this.SetAllNull();

            if (ri.ItemType != null) this.ItemType = ri.ItemType.Value;
            if (ri.RepeatedItemCount != null) this.RepeatedItemCount = ri.RepeatedItemCount.Value;
            if (ri.Index != null) this.Index = ri.Index.Value;

            MemberPropertyIndex mpi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(ri))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(MemberPropertyIndex))
                    {
                        mpi = (MemberPropertyIndex)oxr.LoadCurrentElement();
                        if (mpi.Val != null) this.MemberPropertyIndexes.Add(mpi.Val.Value);
                        else this.MemberPropertyIndexes.Add(0);
                    }
                }
            }
        }

        internal RowItem ToRowItem()
        {
            RowItem ri = new RowItem();
            if (this.ItemType != ItemValues.Data) ri.ItemType = this.ItemType;
            if (this.RepeatedItemCount != 0) ri.RepeatedItemCount = this.RepeatedItemCount;
            if (this.Index != 0) ri.Index = this.Index;

            foreach (int i in this.MemberPropertyIndexes)
            {
                if (i != 0) ri.Append(new MemberPropertyIndex() { Val = i });
                else ri.Append(new MemberPropertyIndex());
            }

            return ri;
        }

        internal SLRowItem Clone()
        {
            SLRowItem ri = new SLRowItem();
            ri.ItemType = this.ItemType;
            ri.RepeatedItemCount = this.RepeatedItemCount;
            ri.Index = this.Index;

            ri.MemberPropertyIndexes = new List<int>();
            foreach (int i in this.MemberPropertyIndexes)
            {
                ri.MemberPropertyIndexes.Add(i);
            }

            return ri;
        }
    }
}
