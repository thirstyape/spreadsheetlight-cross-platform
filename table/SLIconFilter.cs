using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLIconFilter
    {
        internal IconSetValues IconSet { get; set; }
        internal uint? IconId { get; set; }

        internal SLIconFilter()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.IconSet = IconSetValues.ThreeArrows;
            this.IconId = null;
        }

        internal void FromIconFilter(IconFilter icf)
        {
            this.SetAllNull();

            this.IconSet = icf.IconSet.Value;
            if (icf.IconId != null) this.IconId = icf.IconId.Value;
        }

        internal IconFilter ToIconFilter()
        {
            IconFilter icf = new IconFilter();
            icf.IconSet = this.IconSet;
            if (this.IconId != null) icf.IconId = this.IconId.Value;

            return icf;
        }

        internal SLIconFilter Clone()
        {
            SLIconFilter icf = new SLIconFilter();
            icf.IconSet = this.IconSet;
            icf.IconId = this.IconId;

            return icf;
        }
    }
}
