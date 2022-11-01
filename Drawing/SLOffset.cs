using System;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing
{
    internal class SLOffset
    {
        internal long X { get; set; }
        internal long Y { get; set; }

        internal SLOffset()
        {
            this.X = 0;
            this.Y = 0;
        }

        internal A.Offset ToOffset()
        {
            A.Offset off = new A.Offset();
            off.X = this.X;
            off.Y = this.Y;

            return off;
        }

        internal SLOffset Clone()
        {
            SLOffset off = new SLOffset();
            off.X = this.X;
            off.Y = this.Y;

            return off;
        }
    }
}
