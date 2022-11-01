using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLRangeProperties
    {
        internal bool AutoStart { get; set; }
        internal bool AutoEnd { get; set; }
        internal GroupByValues GroupBy { get; set; }
        internal double? StartNumber { get; set; }
        internal double? EndNum { get; set; }
        internal DateTime? StartDate { get; set; }
        internal DateTime? EndDate { get; set; }
        internal double GroupInterval { get; set; }

        internal SLRangeProperties()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.AutoStart = true;
            this.AutoEnd = true;
            this.GroupBy = GroupByValues.Range;
            this.StartNumber = null;
            this.EndNum = null;
            this.StartDate = null;
            this.EndDate = null;
            this.GroupInterval = 1;
        }

        internal void FromRangeProperties(RangeProperties rp)
        {
            this.SetAllNull();

            if (rp.AutoStart != null) this.AutoStart = rp.AutoStart.Value;
            if (rp.AutoEnd != null) this.AutoEnd = rp.AutoEnd.Value;
            if (rp.GroupBy != null) this.GroupBy = rp.GroupBy.Value;
            if (rp.StartNumber != null) this.StartNumber = rp.StartNumber.Value;
            if (rp.EndNum != null) this.EndNum = rp.EndNum.Value;
            if (rp.StartDate != null) this.StartDate = rp.StartDate.Value;
            if (rp.EndDate != null) this.EndDate = rp.EndDate.Value;
            if (rp.GroupInterval != null) this.GroupInterval = rp.GroupInterval.Value;
        }

        internal RangeProperties ToRangeProperties()
        {
            RangeProperties rp = new RangeProperties();
            if (this.AutoStart != true) rp.AutoStart = this.AutoStart;
            if (this.AutoEnd != true) rp.AutoEnd = this.AutoEnd;
            if (this.GroupBy != GroupByValues.Range) rp.GroupBy = this.GroupBy;
            if (this.StartNumber != null) rp.StartNumber = this.StartNumber.Value;
            if (this.EndNum != null) rp.EndNum = this.EndNum.Value;
            if (this.StartDate != null) rp.StartDate = this.StartDate.Value;
            if (this.EndDate != null) rp.EndDate = this.EndDate.Value;
            if (this.GroupInterval != 1) rp.GroupInterval = this.GroupInterval;

            return rp;
        }

        internal SLRangeProperties Clone()
        {
            SLRangeProperties rp = new SLRangeProperties();
            rp.AutoStart = this.AutoStart;
            rp.AutoEnd = this.AutoEnd;
            rp.GroupBy = this.GroupBy;
            rp.StartNumber = this.StartNumber;
            rp.EndNum = this.EndNum;
            rp.StartDate = this.StartDate;
            rp.EndDate = this.EndDate;
            rp.GroupInterval = this.GroupInterval;

            return rp;
        }
    }
}
