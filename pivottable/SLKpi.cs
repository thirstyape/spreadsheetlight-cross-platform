using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLKpi
    {
        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal string DisplayFolder { get; set; }
        internal string MeasureGroup { get; set; }
        internal string ParentKpi { get; set; }
        internal string Value { get; set; }
        internal string Goal { get; set; }
        internal string Status { get; set; }
        internal string Trend { get; set; }
        internal string Weight { get; set; }
        // what happened to Time attribute?

        internal SLKpi()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.UniqueName = "";
            this.Caption = "";
            this.DisplayFolder = "";
            this.MeasureGroup = "";
            this.ParentKpi = "";
            this.Value = "";
            this.Goal = "";
            this.Status = "";
            this.Trend = "";
            this.Weight = "";
        }

        internal void FromKpi(Kpi k)
        {
            this.SetAllNull();

            if (k.UniqueName != null) this.UniqueName = k.UniqueName.Value;
            if (k.Caption != null) this.Caption = k.Caption.Value;
            if (k.DisplayFolder != null) this.DisplayFolder = k.DisplayFolder.Value;
            if (k.MeasureGroup != null) this.MeasureGroup = k.MeasureGroup.Value;
            if (k.ParentKpi != null) this.ParentKpi = k.ParentKpi.Value;
            if (k.Value != null) this.Value = k.Value.Value;
            if (k.Goal != null) this.Goal = k.Goal.Value;
            if (k.Status != null) this.Status = k.Status.Value;
            if (k.Trend != null) this.Trend = k.Trend.Value;
            if (k.Weight != null) this.Weight = k.Weight.Value;
        }

        internal Kpi ToKpi()
        {
            Kpi k = new Kpi();
            k.UniqueName = this.UniqueName;
            if (this.Caption != null && this.Caption.Length > 0) k.Caption = this.Caption;
            if (this.DisplayFolder != null && this.DisplayFolder.Length > 0) k.DisplayFolder = this.DisplayFolder;
            if (this.MeasureGroup != null && this.MeasureGroup.Length > 0) k.MeasureGroup = this.MeasureGroup;
            if (this.ParentKpi != null && this.ParentKpi.Length > 0) k.ParentKpi = this.ParentKpi;
            k.Value = this.Value;
            if (this.Goal != null && this.Goal.Length > 0) k.Goal = this.Goal;
            if (this.Status != null && this.Status.Length > 0) k.Status = this.Status;
            if (this.Trend != null && this.Trend.Length > 0) k.Trend = this.Trend;
            if (this.Weight != null && this.Weight.Length > 0) k.Weight = this.Weight;

            return k;
        }

        internal SLKpi Clone()
        {
            SLKpi k = new SLKpi();
            k.UniqueName = this.UniqueName;
            k.Caption = this.Caption;
            k.DisplayFolder = this.DisplayFolder;
            k.MeasureGroup = this.MeasureGroup;
            k.ParentKpi = this.ParentKpi;
            k.Value = this.Value;
            k.Goal = this.Goal;
            k.Status = this.Status;
            k.Trend = this.Trend;
            k.Weight = this.Weight;

            return k;
        }
    }
}
