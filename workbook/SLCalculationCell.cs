using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLCalculationCell
    {
        internal int RowIndex { get; set; }
        internal int ColumnIndex { get; set; }
        internal int SheetId { get; set; }
        internal bool? InChildChain { get; set; }
        internal bool? NewLevel { get; set; }
        internal bool? NewThread { get; set; }
        internal bool? Array { get; set; }

        internal SLCalculationCell()
        {
            this.SetAllNull();
        }

        internal SLCalculationCell(string CellReference)
        {
            this.SetAllNull();

            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                this.RowIndex = iRowIndex;
                this.ColumnIndex = iColumnIndex;
            }
        }

        private void SetAllNull()
        {
            this.RowIndex = 1;
            this.ColumnIndex = 1;
            this.SheetId = 0;
            this.InChildChain = null;
            this.NewLevel = null;
            this.NewThread = null;
            this.Array = null;
        }

        internal void FromCalculationCell(CalculationCell cc)
        {
            this.SetAllNull();

            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(cc.CellReference.Value, out iRowIndex, out iColumnIndex))
            {
                this.RowIndex = iRowIndex;
                this.ColumnIndex = iColumnIndex;
            }


            this.SheetId = cc.SheetId ?? 0;
            if (cc.InChildChain != null) this.InChildChain = cc.InChildChain.Value;
            if (cc.NewLevel != null) this.NewLevel = cc.NewLevel.Value;
            if (cc.NewThread != null) this.NewThread = cc.NewThread.Value;
            if (cc.Array != null) this.Array = cc.Array.Value;
        }

        internal CalculationCell ToCalculationCell()
        {
            CalculationCell cc = new CalculationCell();
            cc.CellReference = SLTool.ToCellReference(this.RowIndex, this.ColumnIndex);
            cc.SheetId = this.SheetId;
            if (this.InChildChain != null && this.InChildChain.Value) cc.InChildChain = this.InChildChain.Value;
            if (this.NewLevel != null && this.NewLevel.Value) cc.NewLevel = this.NewLevel.Value;
            if (this.NewThread != null && this.NewThread.Value) cc.NewThread = this.NewThread.Value;
            if (this.Array != null && this.Array.Value) cc.Array = this.Array.Value;

            return cc;
        }

        internal SLCalculationCell Clone()
        {
            SLCalculationCell cc = new SLCalculationCell();
            cc.RowIndex = this.RowIndex;
            cc.ColumnIndex = this.ColumnIndex;
            cc.SheetId = this.SheetId;
            cc.InChildChain = this.InChildChain;
            cc.NewLevel = this.NewLevel;
            cc.NewThread = this.NewThread;
            cc.Array = this.Array;

            return cc;
        }
    }
}
