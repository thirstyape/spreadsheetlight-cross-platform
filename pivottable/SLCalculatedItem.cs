using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLCalculatedItem
    {
        internal SLPivotArea PivotArea { get; set; }

        internal uint? Field { get; set; }
        internal string Formula { get; set; }

        internal SLCalculatedItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.PivotArea = new SLPivotArea();
            this.Field = null;
            this.Formula = "";
        }

        internal void FromCalculatedItem(CalculatedItem ci)
        {
            this.SetAllNull();

            if (ci.Field != null) this.Field = ci.Field.Value;
            if (ci.Formula != null) this.Formula = ci.Formula.Value;

            if (ci.PivotArea != null) this.PivotArea.FromPivotArea(ci.PivotArea);
        }

        internal CalculatedItem ToCalculatedItem()
        {
            CalculatedItem ci = new CalculatedItem();
            if (this.Field != null) ci.Field = this.Field.Value;
            if (this.Formula != null && this.Formula.Length > 0) ci.Formula = this.Formula;

            ci.PivotArea = this.PivotArea.ToPivotArea();

            return ci;
        }

        internal SLCalculatedItem Clone()
        {
            SLCalculatedItem ci = new SLCalculatedItem();
            ci.Field = this.Field;
            ci.Formula = this.Formula;
            ci.PivotArea = this.PivotArea.Clone();

            return ci;
        }
    }
}
