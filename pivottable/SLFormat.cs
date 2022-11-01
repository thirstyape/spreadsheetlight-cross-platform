using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLFormat
    {
        internal SLPivotArea PivotArea { get; set; }
        internal FormatActionValues Action { get; set; }
        internal uint? FormatId { get; set; }

        internal SLFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.PivotArea = new SLPivotArea();
            this.Action = FormatActionValues.Formatting;
            this.FormatId = null;
        }

        internal void FromFormat(Format f)
        {
            this.SetAllNull();

            if (f.PivotArea != null) this.PivotArea.FromPivotArea(f.PivotArea);

            if (f.Action != null) this.Action = f.Action.Value;
            if (f.FormatId != null) this.FormatId = f.FormatId.Value;
        }

        internal Format ToFormat()
        {
            Format f = new Format();
            f.PivotArea = this.PivotArea.ToPivotArea();

            if (this.Action != FormatActionValues.Formatting) f.Action = this.Action;
            if (this.FormatId != null) f.FormatId = this.FormatId.Value;

            return f;
        }

        internal SLFormat Clone()
        {
            SLFormat f = new SLFormat();
            f.PivotArea = this.PivotArea.Clone();

            f.Action = this.Action;
            f.FormatId = this.FormatId;

            return f;
        }
    }
}
