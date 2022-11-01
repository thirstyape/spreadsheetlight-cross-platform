using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLMeasureDimensionMap
    {
        internal uint? MeasureGroup { get; set; }
        internal uint? Dimension { get; set; }

        internal SLMeasureDimensionMap()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.MeasureGroup = null;
            this.Dimension = null;
        }

        internal void FromMeasureDimensionMap(MeasureDimensionMap mdm)
        {
            this.SetAllNull();

            if (mdm.MeasureGroup != null) this.MeasureGroup = mdm.MeasureGroup.Value;
            if (mdm.Dimension != null) this.Dimension = mdm.Dimension.Value;
        }

        internal MeasureDimensionMap ToMeasureDimensionMap()
        {
            MeasureDimensionMap mdm = new MeasureDimensionMap();
            if (this.MeasureGroup != null) mdm.MeasureGroup = this.MeasureGroup.Value;
            if (this.Dimension != null) mdm.Dimension = this.Dimension.Value;

            return mdm;
        }

        internal SLMeasureDimensionMap Clone()
        {
            SLMeasureDimensionMap mdm = new SLMeasureDimensionMap();
            mdm.MeasureGroup = this.MeasureGroup;
            mdm.Dimension = this.Dimension;

            return mdm;
        }
    }
}
