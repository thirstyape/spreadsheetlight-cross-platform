using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    internal class SLStringReference
    {
        internal string WorksheetName { get; set; }
        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal string Formula { get; set; }

        // this is StringCache
        internal uint PointCount { get; set; }
        /// <summary>
        /// This takes the place of StringCache
        /// </summary>
        internal List<SLStringPoint> Points { get; set; }

        internal SLStringReference()
        {
            this.WorksheetName = string.Empty;
            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;

            this.Formula = string.Empty;
            this.PointCount = 0;
            this.Points = new List<SLStringPoint>();
        }

        internal C.StringReference ToStringReference()
        {
            C.StringReference sr = new C.StringReference();
            sr.Formula = new C.Formula(this.Formula);
            sr.StringCache = new C.StringCache();
            sr.StringCache.PointCount = new C.PointCount() { Val = this.PointCount };
            for (int i = 0; i < this.Points.Count; ++i)
            {
                sr.StringCache.Append(this.Points[i].ToStringPoint());
            }

            return sr;
        }

        internal void RefreshFormula()
        {
            this.Formula = SLChartTool.GetChartReferenceFormula(this.WorksheetName, this.StartRowIndex, this.StartColumnIndex, this.EndRowIndex, this.EndColumnIndex);
        }

        internal SLStringReference Clone()
        {
            SLStringReference sr = new SLStringReference();
            sr.WorksheetName = this.WorksheetName;
            sr.StartRowIndex = this.StartRowIndex;
            sr.StartColumnIndex = this.StartColumnIndex;
            sr.EndRowIndex = this.EndRowIndex;
            sr.EndColumnIndex = this.EndColumnIndex;
            sr.Formula = this.Formula;
            sr.PointCount = this.PointCount;
            for (int i = 0; i < this.Points.Count; ++i)
            {
                sr.Points.Add(this.Points[i].Clone());
            }

            return sr;
        }
    }
}
