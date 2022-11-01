using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLChartFormat
    {
        internal SLPivotArea PivotArea { get; set; }
        internal uint Chart { get; set; }
        internal uint Format { get; set; }
        internal bool Series { get; set; }

        internal SLChartFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.PivotArea = new SLPivotArea();
            this.Chart = 0;
            this.Format = 0;
            this.Series = false;
        }

        internal void FromChartFormat(ChartFormat cf)
        {
            this.SetAllNull();

            if (cf.PivotArea != null) this.PivotArea.FromPivotArea(cf.PivotArea);

            if (cf.Chart != null) this.Chart = cf.Chart.Value;
            if (cf.Format != null) this.Format = cf.Format.Value;
            if (cf.Series != null) this.Series = cf.Series.Value;
        }

        internal ChartFormat ToChartFormat()
        {
            ChartFormat cf = new ChartFormat();
            cf.PivotArea = this.PivotArea.ToPivotArea();

            cf.Chart = this.Chart;
            cf.Format = this.Format;
            if (this.Series != false) cf.Series = this.Series;

            return cf;
        }

        internal SLChartFormat Clone()
        {
            SLChartFormat cf = new SLChartFormat();
            cf.PivotArea = this.PivotArea.Clone();
            cf.Chart = this.Chart;
            cf.Format = this.Format;
            cf.Series = this.Series;

            return cf;
        }
    }
}
