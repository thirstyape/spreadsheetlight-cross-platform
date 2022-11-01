using System;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing
{
    internal class SLExtents
    {
        internal long Cx { get; set; }
        internal long Cy { get; set; }

        internal SLExtents()
        {
            this.Cx = 0;
            this.Cy = 0;
        }

        internal A.Extents ToExtents()
        {
            A.Extents ext = new A.Extents();
            ext.Cx = this.Cx;
            ext.Cy = this.Cy;

            return ext;
        }

        internal SLExtents Clone()
        {
            SLExtents ext = new SLExtents();
            ext.Cx = this.Cx;
            ext.Cy = this.Cy;

            return ext;
        }
    }
}
