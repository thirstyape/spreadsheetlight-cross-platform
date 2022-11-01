using System;
using System.Collections.Generic;

namespace SpreadsheetLight
{
    internal struct SLSortItem
    {
        internal double Number;
        internal string Text;
        internal int Index;

        internal SLSortItem(double SortNumber, string SortText, int SortIndex)
        {
            Number = SortNumber;
            Text = SortText;
            Index = SortIndex;
        }
    }

    internal class SLSortItemNumberComparer : IComparer<SLSortItem>
    {
        public int Compare(SLSortItem si1, SLSortItem si2)
        {
            return si1.Number.CompareTo(si2.Number);
        }
    }

    internal class SLSortItemTextComparer : IComparer<SLSortItem>
    {
        public int Compare(SLSortItem si1, SLSortItem si2)
        {
            return si1.Text.CompareTo(si2.Text);
        }
    }
}
