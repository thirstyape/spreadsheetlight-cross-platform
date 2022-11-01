using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLBreak
    {
        internal uint Id { get; set; }
        internal uint Min { get; set; }
        internal uint Max { get; set; }
        internal bool ManualPageBreak { get; set; }
        internal bool PivotTablePageBreak { get; set; }

        internal SLBreak()
        {
            this.SetAllNull();
        }

        internal void SetAllNull()
        {
            this.Id = 0;
            this.Min = 0;
            this.Max = 0;
            this.ManualPageBreak = false;
            this.PivotTablePageBreak = false;
        }

        internal void FromBreak(Break b)
        {
            this.SetAllNull();
            if (b.Id != null) this.Id = b.Id;
            if (b.Min != null) this.Min = b.Min;
            if (b.Max != null) this.Max = b.Max;
            if (b.ManualPageBreak != null) this.ManualPageBreak = b.ManualPageBreak;
            if (b.PivotTablePageBreak != null) this.PivotTablePageBreak = b.PivotTablePageBreak;
        }

        internal Break ToBreak()
        {
            Break b = new Break();
            if (this.Id != 0) b.Id = this.Id;
            if (this.Min != 0) b.Min = this.Min;
            if (this.Max != 0) b.Max = this.Max;
            if (this.ManualPageBreak != false) b.ManualPageBreak = this.ManualPageBreak;
            if (this.PivotTablePageBreak != false) b.PivotTablePageBreak = this.PivotTablePageBreak;

            return b;
        }

        internal SLBreak Clone()
        {
            SLBreak b = new SLBreak();
            b.Id = this.Id;
            b.Min = this.Min;
            b.Max = this.Max;
            b.ManualPageBreak = this.ManualPageBreak;
            b.PivotTablePageBreak = this.PivotTablePageBreak;

            return b;
        }
    }
}
