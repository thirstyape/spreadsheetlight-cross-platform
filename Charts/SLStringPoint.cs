using System;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    internal class SLStringPoint
    {
        internal string NumericValue { get; set; }
        internal uint Index { get; set; }

        internal SLStringPoint()
        {
            this.NumericValue = string.Empty;
            this.Index = 0;
        }

        internal C.StringPoint ToStringPoint()
        {
            C.StringPoint sp = new C.StringPoint();
            sp.Index = this.Index;
            sp.NumericValue = new C.NumericValue(this.NumericValue);

            return sp;
        }

        internal SLStringPoint Clone()
        {
            SLStringPoint sp = new SLStringPoint();
            sp.NumericValue = this.NumericValue;
            sp.Index = this.Index;

            return sp;
        }
    }
}
