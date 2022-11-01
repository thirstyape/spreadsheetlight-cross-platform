using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLPageField
    {
        internal int Field { get; set; }
        internal uint? Item { get; set; }
        internal int Hierarchy { get; set; }
        internal string Name { get; set; }
        internal string Caption { get; set; }

        internal SLPageField()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Field = 0;
            this.Item = null;
            this.Hierarchy = 0;
            this.Name = "";
            this.Caption = "";
        }

        internal void FromPageField(PageField pf)
        {
            this.SetAllNull();

            if (pf.Field != null) this.Field = pf.Field.Value;
            if (pf.Item != null) this.Item = pf.Item.Value;
            if (pf.Hierarchy != null) this.Hierarchy = pf.Hierarchy.Value;
            if (pf.Name != null) this.Name = pf.Name.Value;
            if (pf.Caption != null) this.Caption = pf.Caption.Value;
        }

        internal PageField ToPageField()
        {
            PageField pf = new PageField();
            pf.Field = this.Field;
            if (this.Item != null) pf.Item = this.Item.Value;
            pf.Hierarchy = this.Hierarchy;
            if (this.Name != null && this.Name.Length > 0) pf.Name = this.Name;
            if (this.Caption != null && this.Caption.Length > 0) pf.Caption = this.Caption;

            return pf;
        }

        internal SLPageField Clone()
        {
            SLPageField pf = new SLPageField();
            pf.Field = this.Field;
            pf.Item = this.Item;
            pf.Hierarchy = this.Hierarchy;
            pf.Name = this.Name;
            pf.Caption = this.Caption;

            return pf;
        }
    }
}
