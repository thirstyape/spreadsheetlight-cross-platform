using System;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    internal class SLNumericPoint
    {
        internal string NumericValue { get; set; }
        internal uint Index { get; set; }
        internal string FormatCode { get; set; }

        internal SLNumericPoint()
        {
            this.NumericValue = string.Empty;
            this.Index = 0;
            this.FormatCode = string.Empty;
        }

        internal C.NumericPoint ToNumericPoint()
        {
            C.NumericPoint np = new C.NumericPoint();
            np.Index = this.Index;
            if (this.FormatCode.Length > 0) np.FormatCode = this.FormatCode;
            np.NumericValue = new C.NumericValue(this.NumericValue);

            return np;
        }

        internal SLNumericPoint Clone()
        {
            SLNumericPoint np = new SLNumericPoint();
            np.NumericValue = this.NumericValue;
            np.Index = this.Index;
            np.FormatCode = this.FormatCode;

            return np;
        }
    }
}
