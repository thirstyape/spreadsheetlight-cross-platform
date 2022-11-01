using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLColorFilter
    {
        internal uint? FormatId { get; set; }
        internal bool? CellColor { get; set; }

        internal SLColorFilter()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.FormatId = null;
            this.CellColor = null;
        }

        internal void FromColorFilter(ColorFilter cf)
        {
            this.SetAllNull();

            if (cf.FormatId != null) this.FormatId = cf.FormatId.Value;
            if (cf.CellColor != null && !cf.CellColor.Value) this.CellColor = cf.CellColor.Value;
        }

        internal ColorFilter ToColorFilter()
        {
            ColorFilter cf = new ColorFilter();
            if (this.FormatId != null) cf.FormatId = this.FormatId.Value;
            if (this.CellColor != null && !this.CellColor.Value) cf.CellColor = this.CellColor.Value;

            return cf;
        }

        internal SLColorFilter Clone()
        {
            SLColorFilter cf = new SLColorFilter();
            cf.FormatId = this.FormatId;
            cf.CellColor = this.CellColor;

            return cf;
        }
    }
}
