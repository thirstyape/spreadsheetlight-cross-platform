using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLDataField
    {
        internal string Name { get; set; }
        internal uint Field { get; set; }
        internal DataConsolidateFunctionValues Subtotal { get; set; }
        internal ShowDataAsValues ShowDataAs { get; set; }
        internal int BaseField { get; set; }
        internal uint BaseItem { get; set; }
        internal uint? NumberFormatId { get; set; }

        internal SLDataField()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Name = "";
            this.Field = 1;
            this.Subtotal = DataConsolidateFunctionValues.Sum;
            this.ShowDataAs = ShowDataAsValues.Normal;
            this.BaseField = -1;

            // why the weird default value? It's 2^20 + 2^8 for what it's worth...
            this.BaseItem = 1048832;

            this.NumberFormatId = null;
        }

        internal void FromDataField(DataField df)
        {
            this.SetAllNull();

            if (df.Name != null) this.Name = df.Name.Value;
            if (df.Field != null) this.Field = df.Field.Value;
            if (df.Subtotal != null) this.Subtotal = df.Subtotal.Value;
            if (df.ShowDataAs != null) this.ShowDataAs = df.ShowDataAs.Value;
            if (df.BaseField != null) this.BaseField = df.BaseField.Value;
            if (df.BaseItem != null) this.BaseItem = df.BaseItem.Value;
            if (df.NumberFormatId != null) this.NumberFormatId = df.NumberFormatId.Value;
        }

        internal DataField ToDataField()
        {
            DataField df = new DataField();
            if (this.Name != null && this.Name.Length > 0) df.Name = this.Name;
            df.Field = this.Field;
            if (this.Subtotal != DataConsolidateFunctionValues.Sum) df.Subtotal = this.Subtotal;
            if (this.ShowDataAs != ShowDataAsValues.Normal) df.ShowDataAs = this.ShowDataAs;
            if (this.BaseField != -1) df.BaseField = this.BaseField;
            if (this.BaseItem != 1048832) df.BaseItem = this.BaseItem;
            if (this.NumberFormatId != null) df.NumberFormatId = this.NumberFormatId.Value;

            return df;
        }

        internal SLDataField Clone()
        {
            SLDataField df = new SLDataField();
            df.Name = this.Name;
            df.Field = this.Field;
            df.Subtotal = this.Subtotal;
            df.ShowDataAs = this.ShowDataAs;
            df.BaseField = this.BaseField;
            df.BaseItem = this.BaseItem;
            df.NumberFormatId = this.NumberFormatId;

            return df;
        }
    }
}
