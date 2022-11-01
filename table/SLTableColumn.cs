using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLTableColumn
    {
        internal bool HasCalculatedColumnFormula;
        internal SLCalculatedColumnFormula CalculatedColumnFormula { get; set; }

        internal bool HasTotalsRowFormula;
        internal SLTotalsRowFormula TotalsRowFormula { get; set; }

        internal bool HasXmlColumnProperties;
        internal SLXmlColumnProperties XmlColumnProperties { get; set; }

        internal uint Id { get; set; }
        internal string UniqueName { get; set; }
        internal string Name { get; set; }

        internal bool HasTotalsRowFunction;
        private TotalsRowFunctionValues vTotalsRowFunction;
        internal TotalsRowFunctionValues TotalsRowFunction
        {
            get { return vTotalsRowFunction; }
            set
            {
                vTotalsRowFunction = value;
                HasTotalsRowFunction = vTotalsRowFunction != TotalsRowFunctionValues.None ? true : false;
            }
        }

        internal string TotalsRowLabel { get; set; }
        internal uint? QueryTableFieldId { get; set; }
        internal uint? HeaderRowDifferentialFormattingId { get; set; }
        internal uint? DataFormatId { get; set; }
        internal uint? TotalsRowDifferentialFormattingId { get; set; }
        internal string HeaderRowCellStyle { get; set; }
        internal string DataCellStyle { get; set; }
        internal string TotalsRowCellStyle { get; set; }

        internal SLTableColumn()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.CalculatedColumnFormula = new SLCalculatedColumnFormula();
            this.HasCalculatedColumnFormula = false;
            this.TotalsRowFormula = new SLTotalsRowFormula();
            this.HasTotalsRowFormula = false;
            this.XmlColumnProperties = new SLXmlColumnProperties();
            this.HasXmlColumnProperties = false;

            this.Id = 0;
            this.UniqueName = null;
            this.Name = string.Empty;
            this.TotalsRowFunction = TotalsRowFunctionValues.None;
            this.HasTotalsRowFunction = false;
            this.TotalsRowLabel = null;
            this.QueryTableFieldId = null;
            this.HeaderRowDifferentialFormattingId = null;
            this.DataFormatId = null;
            this.TotalsRowDifferentialFormattingId = null;
            this.HeaderRowCellStyle = null;
            this.DataCellStyle = null;
            this.TotalsRowCellStyle = null;
        }

        internal void FromTableColumn(TableColumn tc)
        {
            this.SetAllNull();

            if (tc.CalculatedColumnFormula != null)
            {
                this.HasCalculatedColumnFormula = true;
                this.CalculatedColumnFormula.FromCalculatedColumnFormula(tc.CalculatedColumnFormula);
            }
            if (tc.TotalsRowFormula != null)
            {
                this.HasTotalsRowFormula = true;
                this.TotalsRowFormula.FromTotalsRowFormula(tc.TotalsRowFormula);
            }
            if (tc.XmlColumnProperties != null)
            {
                this.HasXmlColumnProperties = true;
                this.XmlColumnProperties.FromXmlColumnProperties(tc.XmlColumnProperties);
            }

            this.Id = tc.Id.Value;
            if (tc.UniqueName != null) this.UniqueName = tc.UniqueName.Value;
            this.Name = tc.Name.Value;

            if (tc.TotalsRowFunction != null) this.TotalsRowFunction = tc.TotalsRowFunction.Value;
            if (tc.TotalsRowLabel != null) this.TotalsRowLabel = tc.TotalsRowLabel.Value;
            if (tc.QueryTableFieldId != null) this.QueryTableFieldId = tc.QueryTableFieldId.Value;
            if (tc.HeaderRowDifferentialFormattingId != null) this.HeaderRowDifferentialFormattingId = tc.HeaderRowDifferentialFormattingId.Value;
            if (tc.DataFormatId != null) this.DataFormatId = tc.DataFormatId.Value;
            if (tc.TotalsRowDifferentialFormattingId != null) this.TotalsRowDifferentialFormattingId = tc.TotalsRowDifferentialFormattingId.Value;
            if (tc.HeaderRowCellStyle != null) this.HeaderRowCellStyle = tc.HeaderRowCellStyle.Value;
            if (tc.DataCellStyle != null) this.DataCellStyle = tc.DataCellStyle.Value;
            if (tc.TotalsRowCellStyle != null) this.TotalsRowCellStyle = tc.TotalsRowCellStyle.Value;
        }

        internal TableColumn ToTableColumn()
        {
            TableColumn tc = new TableColumn();
            if (HasCalculatedColumnFormula)
            {
                tc.CalculatedColumnFormula = this.CalculatedColumnFormula.ToCalculatedColumnFormula();
            }
            if (HasTotalsRowFormula)
            {
                tc.TotalsRowFormula = this.TotalsRowFormula.ToTotalsRowFormula();
            }
            if (HasXmlColumnProperties)
            {
                tc.XmlColumnProperties = this.XmlColumnProperties.ToXmlColumnProperties();
            }

            tc.Id = this.Id;
            if (this.UniqueName != null) tc.UniqueName = this.UniqueName;
            tc.Name = this.Name;

            if (HasTotalsRowFunction) tc.TotalsRowFunction = this.TotalsRowFunction;
            if (this.TotalsRowLabel != null) tc.TotalsRowLabel = this.TotalsRowLabel;
            if (this.QueryTableFieldId != null) tc.QueryTableFieldId = this.QueryTableFieldId.Value;
            if (this.HeaderRowDifferentialFormattingId != null) tc.HeaderRowDifferentialFormattingId = this.HeaderRowDifferentialFormattingId.Value;
            if (this.DataFormatId != null) tc.DataFormatId = this.DataFormatId.Value;
            if (this.TotalsRowDifferentialFormattingId != null) tc.TotalsRowDifferentialFormattingId = this.TotalsRowDifferentialFormattingId.Value;
            if (this.HeaderRowCellStyle != null) tc.HeaderRowCellStyle = this.HeaderRowCellStyle;
            if (this.DataCellStyle != null) tc.DataCellStyle = this.DataCellStyle;
            if (this.TotalsRowCellStyle != null) tc.TotalsRowCellStyle = this.TotalsRowCellStyle;

            return tc;
        }

        internal SLTableColumn Clone()
        {
            SLTableColumn tc = new SLTableColumn();
            tc.HasCalculatedColumnFormula = this.HasCalculatedColumnFormula;
            tc.CalculatedColumnFormula = this.CalculatedColumnFormula.Clone();
            tc.HasTotalsRowFormula = this.HasTotalsRowFormula;
            tc.TotalsRowFormula = this.TotalsRowFormula.Clone();
            tc.HasXmlColumnProperties = this.HasXmlColumnProperties;
            tc.XmlColumnProperties = this.XmlColumnProperties.Clone();
            tc.Id = this.Id;
            tc.UniqueName = this.UniqueName;
            tc.Name = this.Name;
            tc.HasTotalsRowFunction = this.HasTotalsRowFunction;
            tc.vTotalsRowFunction = this.vTotalsRowFunction;
            tc.TotalsRowLabel = this.TotalsRowLabel;
            tc.QueryTableFieldId = this.QueryTableFieldId;
            tc.HeaderRowDifferentialFormattingId = this.HeaderRowDifferentialFormattingId;
            tc.DataFormatId = this.DataFormatId;
            tc.TotalsRowDifferentialFormattingId = this.TotalsRowDifferentialFormattingId;
            tc.HeaderRowCellStyle = this.HeaderRowCellStyle;
            tc.DataCellStyle = this.DataCellStyle;
            tc.TotalsRowCellStyle = this.TotalsRowCellStyle;

            return tc;
        }
    }
}
