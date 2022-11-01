using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLDateGroupItem
    {
        internal ushort Year { get; set; }
        internal ushort? Month { get; set; }
        internal ushort? Day { get; set; }
        internal ushort? Hour { get; set; }
        internal ushort? Minute { get; set; }
        internal ushort? Second { get; set; }
        internal DateTimeGroupingValues DateTimeGrouping { get; set; }

        internal SLDateGroupItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Year = (ushort)DateTime.Now.Year;
            this.Month = null;
            this.Day = null;
            this.Hour = null;
            this.Minute = null;
            this.Second = null;
            this.DateTimeGrouping = DateTimeGroupingValues.Year;
        }

        internal void FromDateGroupItem(DateGroupItem dgi)
        {
            this.SetAllNull();

            this.Year = dgi.Year.Value;
            if (dgi.Month != null) this.Month = dgi.Month.Value;
            if (dgi.Day != null) this.Day = dgi.Day.Value;
            if (dgi.Hour != null) this.Hour = dgi.Hour.Value;
            if (dgi.Minute != null) this.Minute = dgi.Minute.Value;
            if (dgi.Second != null) this.Second = dgi.Second.Value;
            this.DateTimeGrouping = dgi.DateTimeGrouping.Value;
        }

        internal DateGroupItem ToDateGroupItem()
        {
            DateGroupItem dgi = new DateGroupItem();
            dgi.Year = this.Year;
            if (this.Month != null) dgi.Month = this.Month.Value;
            if (this.Day != null) dgi.Day = this.Day.Value;
            if (this.Hour != null) dgi.Hour = this.Hour.Value;
            if (this.Minute != null) dgi.Minute = this.Minute.Value;
            if (this.Second != null) dgi.Second = this.Second.Value;
            dgi.DateTimeGrouping = this.DateTimeGrouping;

            return dgi;
        }

        internal SLDateGroupItem Clone()
        {
            SLDateGroupItem dgi = new SLDateGroupItem();
            dgi.Year = this.Year;
            dgi.Month = this.Month;
            dgi.Day = this.Day;
            dgi.Hour = this.Hour;
            dgi.Minute = this.Minute;
            dgi.Second = this.Second;
            dgi.DateTimeGrouping = this.DateTimeGrouping;

            return dgi;
        }
    }
}
