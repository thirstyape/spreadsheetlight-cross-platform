using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLDynamicFilter
    {
        internal DynamicFilterValues Type { get; set; }
        internal double? Val { get; set; }
        internal double? MaxVal { get; set; }

        internal SLDynamicFilter()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Type = DynamicFilterValues.Null;
            this.Val = null;
            this.MaxVal = null;
        }

        internal void FromDynamicFilter(DynamicFilter df)
        {
            this.SetAllNull();

            this.Type = df.Type.Value;
            if (df.Val != null) this.Val = df.Val.Value;
            if (df.MaxVal != null) this.MaxVal = df.MaxVal.Value;
        }

        internal DynamicFilter ToDynamicFilter()
        {
            DynamicFilter df = new DynamicFilter();
            df.Type = this.Type;
            if (this.Val != null) df.Val = this.Val.Value;
            if (this.MaxVal != null) df.MaxVal = this.MaxVal.Value;

            return df;
        }

        internal SLDynamicFilter Clone()
        {
            SLDynamicFilter df = new SLDynamicFilter();
            df.Type = this.Type;
            df.Val = this.Val;
            df.MaxVal = this.MaxVal;

            return df;
        }
    }
}
