using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLFilter
    {
        internal string Val { get; set; }

        internal SLFilter()
        {
            this.Val = string.Empty;
        }

        internal void FromFilter(Filter f)
        {
            this.Val = f.Val ?? string.Empty;
        }

        internal Filter ToFilter()
        {
            Filter f = new Filter();
            f.Val = this.Val;

            return f;
        }

        internal SLFilter Clone()
        {
            SLFilter f = new SLFilter();
            f.Val = this.Val;

            return f;
        }
    }
}
