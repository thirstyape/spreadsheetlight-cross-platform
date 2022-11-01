using System;

namespace SpreadsheetLight
{
    internal struct SLRowColumnStyleHistory
    {
        internal bool IsRow;
        internal int Index;

        internal SLRowColumnStyleHistory(bool IsRow, int Index)
        {
            this.IsRow = IsRow;
            this.Index = Index;
        }
    }
}
