using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for shared cell formula.
    /// </summary>
    public class SLSharedCellFormula
    {
        /// <summary>
        /// The base row index for the shared cell formula.
        /// </summary>
        public int BaseCellRowIndex { get; set; }

        /// <summary>
        /// The base column index for the shared cell formula.
        /// </summary>
        public int BaseCellColumnIndex { get; set; }

        /// <summary>
        /// The shared index.
        /// </summary>
        public uint SharedIndex { get; set; }

        /// <summary>
        /// The cell range(s) under the shared cell formula.
        /// </summary>
        public List<SLCellPointRange> Reference { get; set; }

        /// <summary>
        /// The shared cell formula text.
        /// </summary>
        public string FormulaText { get; set; }

        /// <summary>
        /// Initializes an instance of SLSharedCellFormula.
        /// </summary>
        public SLSharedCellFormula()
        {
            this.BaseCellRowIndex = -1;
            this.BaseCellColumnIndex = -1;
            this.SharedIndex = 0;
            this.Reference = new List<SLCellPointRange>();
            this.FormulaText = string.Empty;
        }

        /// <summary>
        /// Clone a new instance of SLSharedCellFormula.
        /// </summary>
        /// <returns>A cloned SLSharedCellFormula object.</returns>
        public SLSharedCellFormula Clone()
        {
            SLSharedCellFormula scf = new SLSharedCellFormula();
            scf.BaseCellRowIndex = this.BaseCellRowIndex;
            scf.BaseCellColumnIndex = this.BaseCellColumnIndex;
            scf.SharedIndex = this.SharedIndex;

            for (int i = 0; i < this.Reference.Count; ++i)
            {
                scf.Reference.Add(new SLCellPointRange(
                    this.Reference[i].StartRowIndex,
                    this.Reference[i].StartColumnIndex,
                    this.Reference[i].EndRowIndex,
                    this.Reference[i].EndColumnIndex));
            }

            scf.FormulaText = this.FormulaText;

            return scf;
        }
    }
}
