using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLTotalsRowFormula
    {
        internal bool Array { get; set; }
        internal string Text { get; set; }

        internal SLTotalsRowFormula()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Array = false;
            this.Text = string.Empty;
        }

        internal void FromTotalsRowFormula(TotalsRowFormula trf)
        {
            this.SetAllNull();

            if (trf.Array != null && trf.Array.Value) this.Array = true;
            this.Text = trf.Text;
        }

        internal TotalsRowFormula ToTotalsRowFormula()
        {
            TotalsRowFormula trf = new TotalsRowFormula();
            if (this.Array) trf.Array = this.Array;

            if (SLTool.ToPreserveSpace(this.Text))
            {
                trf.Space = SpaceProcessingModeValues.Preserve;
            }
            trf.Text = this.Text;

            return trf;
        }

        internal SLTotalsRowFormula Clone()
        {
            SLTotalsRowFormula trf = new SLTotalsRowFormula();
            trf.Array = this.Array;
            trf.Text = this.Text;

            return trf;
        }
    }
}
