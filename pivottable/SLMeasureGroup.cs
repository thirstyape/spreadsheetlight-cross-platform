using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLMeasureGroup
    {
        internal string Name { get; set; }
        internal string Caption { get; set; }

        internal SLMeasureGroup()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Name = "";
            this.Caption = "";
        }

        internal void FromMeasureGroup(MeasureGroup mg)
        {
            this.SetAllNull();

            if (mg.Name != null) this.Name = mg.Name.Value;
            if (mg.Caption != null) this.Caption = mg.Caption.Value;
        }

        internal MeasureGroup ToMeasureGroup()
        {
            MeasureGroup mg = new MeasureGroup();
            mg.Name = this.Name;
            mg.Caption = this.Caption;

            return mg;
        }

        internal SLMeasureGroup Clone()
        {
            SLMeasureGroup mg = new SLMeasureGroup();
            mg.Name = this.Name;
            mg.Caption = this.Caption;

            return mg;
        }
    }
}
