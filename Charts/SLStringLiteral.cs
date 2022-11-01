using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    internal class SLStringLiteral
    {
        internal uint PointCount { get; set; }
        internal List<SLStringPoint> Points { get; set; }

        internal SLStringLiteral()
        {
            this.PointCount = 0;
            this.Points = new List<SLStringPoint>();
        }

        internal C.StringLiteral ToStringLiteral()
        {
            C.StringLiteral sl = new C.StringLiteral();
            sl.PointCount = new C.PointCount() { Val = this.PointCount };
            for (int i = 0; i < this.Points.Count; ++i)
            {
                sl.Append(this.Points[i].ToStringPoint());
            }

            return sl;
        }

        internal SLStringLiteral Clone()
        {
            SLStringLiteral sl = new SLStringLiteral();
            sl.PointCount = this.PointCount;
            for (int i = 0; i < this.Points.Count; ++i)
            {
                sl.Points.Add(this.Points[i].Clone());
            }

            return sl;
        }
    }
}
