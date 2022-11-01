using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLAutoSortScope
    {
        internal SLPivotArea PivotArea { get; set; }

        internal SLAutoSortScope()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.PivotArea = new SLPivotArea();
        }

        // ahahahahah... I did *not* just come up with this variable name... :)
        internal void FromAutoSortScope(AutoSortScope ass)
        {
            this.SetAllNull();

            if (ass.PivotArea != null) this.PivotArea.FromPivotArea(ass.PivotArea);
        }

        internal AutoSortScope ToAutoSortScope()
        {
            AutoSortScope ass = new AutoSortScope();
            ass.PivotArea = this.PivotArea.ToPivotArea();

            return ass;
        }

        internal SLAutoSortScope Clone()
        {
            SLAutoSortScope ass = new SLAutoSortScope();
            ass.PivotArea = this.PivotArea.Clone();

            return ass;
        }
    }
}
