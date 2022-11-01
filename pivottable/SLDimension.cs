using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLDimension
    {
        internal bool Measure { get; set; }
        internal string Name { get; set; }
        internal string UniqueName { get; set; }
        internal string Caption { get; set; }

        internal SLDimension()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Measure = false;
            this.Name = "";
            this.UniqueName = "";
            this.Caption = "";
        }

        internal void FromDimension(Dimension d)
        {
            this.SetAllNull();

            if (d.Measure != null) this.Measure = d.Measure.Value;
            if (d.Name != null) this.Name = d.Name.Value;
            if (d.UniqueName != null) this.UniqueName = d.UniqueName.Value;
            if (d.Caption != null) this.Caption = d.Caption.Value;
        }

        internal Dimension ToDimension()
        {
            Dimension d = new Dimension();
            if (this.Measure != false) d.Measure = this.Measure;
            d.Name = this.Name;
            d.UniqueName = this.UniqueName;
            d.Caption = this.Caption;

            return d;
        }

        internal SLDimension Clone()
        {
            SLDimension d = new SLDimension();
            d.Measure = this.Measure;
            d.Name = this.Name;
            d.UniqueName = this.UniqueName;
            d.Caption = this.Caption;

            return d;
        }
    }
}
