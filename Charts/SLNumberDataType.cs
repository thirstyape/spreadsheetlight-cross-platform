using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    internal class SLNumberingCache : SLNumberDataType
    {
        internal SLNumberingCache() { }

        internal SLNumberingCache Clone()
        {
            SLNumberingCache nc = new SLNumberingCache();
            nc.FormatCode = this.FormatCode;
            nc.PointCount = this.PointCount;
            for (int i = 0; i < this.Points.Count; ++i)
            {
                nc.Points.Add(this.Points[i].Clone());
            }

            return nc;
        }
    }

    internal class SLNumberLiteral : SLNumberDataType
    {
        internal SLNumberLiteral() { }

        internal SLNumberLiteral Clone()
        {
            SLNumberLiteral nl = new SLNumberLiteral();
            nl.FormatCode = this.FormatCode;
            nl.PointCount = this.PointCount;
            for (int i = 0; i < this.Points.Count; ++i)
            {
                nl.Points.Add(this.Points[i].Clone());
            }

            return nl;
        }
    }

    /// <summary>
    /// For NumberingCache and NumberLiteral
    /// </summary>
    internal abstract class SLNumberDataType
    {
        internal string FormatCode { get; set; }
        internal uint PointCount { get; set; }
        internal List<SLNumericPoint> Points { get; set; }

        internal SLNumberDataType()
        {
            this.FormatCode = string.Empty;
            this.PointCount = 0;
            this.Points = new List<SLNumericPoint>();
        }

        internal C.NumberingCache ToNumberingCache()
        {
            C.NumberingCache nc = new C.NumberingCache();
            nc.FormatCode = new C.FormatCode(this.FormatCode);
            nc.PointCount = new C.PointCount() { Val = this.PointCount };
            for (int i = 0; i < this.Points.Count; ++i)
            {
                nc.Append(this.Points[i].ToNumericPoint());
            }

            return nc;
        }

        internal C.NumberLiteral ToNumberLiteral()
        {
            C.NumberLiteral nl = new C.NumberLiteral();
            nl.FormatCode = new C.FormatCode(this.FormatCode);
            nl.PointCount = new C.PointCount() { Val = this.PointCount };
            for (int i = 0; i < this.Points.Count; ++i)
            {
                nl.Append(this.Points[i].ToNumericPoint());
            }

            return nl;
        }
    }
}
