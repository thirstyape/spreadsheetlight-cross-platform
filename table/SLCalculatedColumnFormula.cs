using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLCalculatedColumnFormula
    {
        internal bool Array { get; set; }
        internal string Text { get; set; }

        internal SLCalculatedColumnFormula()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Array = false;
            this.Text = string.Empty;
        }

        internal void FromCalculatedColumnFormula(CalculatedColumnFormula ccf)
        {
            this.SetAllNull();

            if (ccf.Array != null && ccf.Array.Value) this.Array = true;
            this.Text = ccf.Text;
        }

        internal CalculatedColumnFormula ToCalculatedColumnFormula()
        {
            CalculatedColumnFormula ccf = new CalculatedColumnFormula();
            if (this.Array) ccf.Array = this.Array;

            if (SLTool.ToPreserveSpace(this.Text))
            {
                ccf.Space = SpaceProcessingModeValues.Preserve;
            }
            ccf.Text = this.Text;

            return ccf;
        }

        internal SLCalculatedColumnFormula Clone()
        {
            SLCalculatedColumnFormula ccf = new SLCalculatedColumnFormula();
            ccf.Array = this.Array;
            ccf.Text = this.Text;

            return ccf;
        }
    }
}
