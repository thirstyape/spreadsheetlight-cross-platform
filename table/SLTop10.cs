using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLTop10
    {
        internal bool? Top { get; set; }
        internal bool? Percent { get; set; }
        internal double Val { get; set; }
        internal double? FilterValue { get; set; }

        internal SLTop10()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Top = null;
            this.Percent = null;
            this.Val = 0.0;
            this.FilterValue = null;
        }

        internal void FromTop10(Top10 t)
        {
            this.SetAllNull();

            if (t.Top != null) this.Top = t.Top.Value;
            if (t.Percent != null) this.Percent = t.Percent.Value;
            this.Val = t.Val.Value;
            if (t.FilterValue != null) this.FilterValue = t.FilterValue.Value;
        }

        internal Top10 ToTop10()
        {
            Top10 t = new Top10();
            if (this.Top != null && !this.Top.Value) t.Top = this.Top.Value;
            if (this.Percent != null && this.Percent.Value) t.Percent = this.Percent.Value;
            t.Val = this.Val;
            if (this.FilterValue != null) t.FilterValue = this.FilterValue.Value;

            return t;
        }

        internal SLTop10 Clone()
        {
            SLTop10 t = new SLTop10();
            t.Top = this.Top;
            t.Percent = this.Percent;
            t.Val = this.Val;
            t.FilterValue = this.FilterValue;

            return t;
        }
    }
}
