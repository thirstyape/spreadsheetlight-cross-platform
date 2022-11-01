using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// This is for information purposes only! This simulates the DocumentFormat.OpenXml.Spreadsheet.CellFormula class.
    /// </summary>
    public class SLCellFormula
    {
        // We're not going to do preserving space. Don't know the full behaviour for
        // excessively spaced formula text...

        /// <summary>
        /// The formula text.
        /// </summary>
        public string FormulaText { get; set; }

        /// <summary>
        /// The type of formula.
        /// </summary>
        public CellFormulaValues FormulaType { get; set; }

        /// <summary>
        /// If true, then formula is an array formula and the entire array is calculated in full.
        /// If false, then individual cells are calculated as needed.
        /// </summary>
        public bool AlwaysCalculateArray { get; set; }

        /// <summary>
        /// Range of cells where the formula is applied.
        /// </summary>
        public List<SLCellPointRange> Reference { get; set; }

        /// <summary>
        /// True for 2-dimensional data table. False otherwise.
        /// </summary>
        public bool DataTable2D { get; set; }

        /// <summary>
        /// If true, then 1-dimensional data table is a row. Otherwise it's a column.
        /// </summary>
        public bool DataTableRow { get; set; }

        /// <summary>
        /// Whether the first input cell for data table is deleted.
        /// </summary>
        public bool Input1Deleted { get; set; }

        /// <summary>
        /// Whether the second input cell for data table is deleted.
        /// </summary>
        public bool Input2Deleted { get; set; }

        /// <summary>
        /// First input cell for data table.
        /// </summary>
        public string R1 { get; set; }

        /// <summary>
        /// Second input cell for data table.
        /// </summary>
        public string R2 { get; set; }

        /// <summary>
        /// Indicates whether this formula needs to be recalculated.
        /// </summary>
        public bool CalculateCell { get; set; }

        /// <summary>
        /// Shared formula index.
        /// </summary>
        public uint? SharedIndex { get; set; }

        /// <summary>
        /// Specifies that this formula assigns a value to a name.
        /// </summary>
        public bool Bx { get; set; }

        internal SLCellFormula()
        {
            this.SetAllNull();
        }

        internal void SetAllNull()
        {
            this.FormulaText = string.Empty;

            this.FormulaType = CellFormulaValues.Normal;
            this.AlwaysCalculateArray = false;
            this.Reference = new List<SLCellPointRange>();
            this.DataTable2D = false;
            this.DataTableRow = false;
            this.Input1Deleted = false;
            this.Input2Deleted = false;
            this.R1 = "";
            this.R2 = "";
            this.CalculateCell = false;
            this.SharedIndex = null;
            this.Bx = false;
        }

        internal void FromCellFormula(CellFormula cf)
        {
            this.SetAllNull();

            this.FormulaText = cf.Text;
            if (cf.FormulaType != null) this.FormulaType = cf.FormulaType.Value;
            if (cf.AlwaysCalculateArray != null) this.AlwaysCalculateArray = cf.AlwaysCalculateArray.Value;
            if (cf.Reference != null) this.Reference = SLTool.TranslateReferenceToCellPointRangeList(cf.Reference.Value);
            if (cf.DataTable2D != null) this.DataTable2D = cf.DataTable2D.Value;
            if (cf.DataTableRow != null) this.DataTableRow = cf.DataTableRow.Value;
            if (cf.Input1Deleted != null) this.Input1Deleted = cf.Input1Deleted.Value;
            if (cf.Input2Deleted != null) this.Input2Deleted = cf.Input2Deleted.Value;
            if (cf.R1 != null) this.R1 = cf.R1.Value;
            if (cf.R2 != null) this.R2 = cf.R2.Value;
            if (cf.CalculateCell != null) this.CalculateCell = cf.CalculateCell.Value;
            if (cf.SharedIndex != null) this.SharedIndex = cf.SharedIndex.Value;
            if (cf.Bx != null) this.Bx = cf.Bx.Value;
        }

        internal CellFormula ToCellFormula()
        {
            CellFormula cf = new CellFormula();
            cf.Text = this.FormulaText;

            if (this.FormulaType != CellFormulaValues.Normal) cf.FormulaType = this.FormulaType;
            if (this.AlwaysCalculateArray != false) cf.AlwaysCalculateArray = this.AlwaysCalculateArray;
            if (this.Reference.Count > 0) cf.Reference = SLTool.TranslateCellPointRangeToRefSeq(this.Reference);
            if (this.DataTable2D != false) cf.DataTable2D = this.DataTable2D;
            if (this.DataTableRow != false) cf.DataTableRow = this.DataTableRow;
            if (this.Input1Deleted != false) cf.Input1Deleted = this.Input1Deleted;
            if (this.Input2Deleted != false) cf.Input2Deleted = this.Input2Deleted;
            if (this.R1.Length > 0) cf.R1 = this.R1;
            if (this.R2.Length > 0) cf.R2 = this.R2;
            if (this.CalculateCell != false) cf.CalculateCell = this.CalculateCell;
            if (this.SharedIndex != null) cf.SharedIndex = this.SharedIndex.Value;
            if (this.Bx != false) cf.Bx = this.Bx;

            return cf;
        }

        internal static string GetFormulaTypeAttribute(CellFormulaValues cfv)
        {
            string result = "normal";
            switch (cfv)
            {
                case CellFormulaValues.Normal:
                    result = "normal";
                    break;
                case CellFormulaValues.Array:
                    result = "array";
                    break;
                case CellFormulaValues.DataTable:
                    result = "dataTable";
                    break;
                case CellFormulaValues.Shared:
                    result = "shared";
                    break;
            }

            return result;
        }

        internal SLCellFormula Clone()
        {
            SLCellFormula cf = new SLCellFormula();
            cf.FormulaText = this.FormulaText;
            cf.FormulaType = this.FormulaType;
            cf.AlwaysCalculateArray = this.AlwaysCalculateArray;

            for (int i = 0; i < this.Reference.Count; ++i)
            {
                cf.Reference.Add(new SLCellPointRange(
                    this.Reference[i].StartRowIndex,
                    this.Reference[i].StartColumnIndex,
                    this.Reference[i].EndRowIndex,
                    this.Reference[i].EndColumnIndex));
            }

            cf.DataTable2D = this.DataTable2D;
            cf.DataTableRow = this.DataTableRow;
            cf.Input1Deleted = this.Input1Deleted;
            cf.Input2Deleted = this.Input2Deleted;
            cf.R1 = this.R1;
            cf.R2 = this.R2;
            cf.CalculateCell = this.CalculateCell;
            cf.SharedIndex = this.SharedIndex;
            cf.Bx = this.Bx;

            return cf;
        }
    }
}
