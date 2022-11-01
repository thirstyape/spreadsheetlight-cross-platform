using System;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    internal class SLNumberReference
    {
        internal string WorksheetName { get; set; }
        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal string Formula { get; set; }
        internal SLNumberingCache NumberingCache { get; set; }

        internal SLNumberReference()
        {
            this.WorksheetName = string.Empty;
            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;

            this.Formula = string.Empty;
            this.NumberingCache = new SLNumberingCache();
        }

        internal C.NumberReference ToNumberReference()
        {
            C.NumberReference nr = new C.NumberReference();
            nr.Formula = new C.Formula(this.Formula);
            nr.NumberingCache = this.NumberingCache.ToNumberingCache();

            return nr;
        }

        internal void RefreshFormula()
        {
            this.Formula = SLChartTool.GetChartReferenceFormula(this.WorksheetName, this.StartRowIndex, this.StartColumnIndex, this.EndRowIndex, this.EndColumnIndex);
        }

        internal SLNumberReference Clone()
        {
            SLNumberReference nr = new SLNumberReference();
            nr.WorksheetName = this.WorksheetName;
            nr.StartRowIndex = this.StartRowIndex;
            nr.StartColumnIndex = this.StartColumnIndex;
            nr.EndRowIndex = this.EndRowIndex;
            nr.EndColumnIndex = this.EndColumnIndex;
            nr.Formula = this.Formula;
            nr.NumberingCache = this.NumberingCache.Clone();

            return nr;
        }
    }
}
