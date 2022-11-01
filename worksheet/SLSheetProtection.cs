using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for specifying worksheet protection. This simulates the DocumentFormat.OpenXml.Spreadsheet.SheetProtection class.
    /// </summary>
    public class SLSheetProtection
    {
        internal string AlgorithmName { get; set; }
        internal string HashValue { get; set; }
        internal string SaltValue { get; set; }
        internal uint? SpinCount { get; set; }
        internal string Password { get; set; }

        internal bool? Sheet { get; set; }

        // all the following properties take the negation of the default boolean value
        // of the corresponding attributes

        internal bool? bAllowEditObjects;
        /// <summary>
        /// Allow editing of objects even if sheet is protected.
        /// </summary>
        public bool AllowEditObjects
        {
            get { return bAllowEditObjects ?? true; }
            set { bAllowEditObjects = value; }
        }

        internal bool? bAllowEditScenarios;
        /// <summary>
        /// Allow editing of scenarios even if sheet is protected.
        /// </summary>
        public bool AllowEditScenarios
        {
            get { return bAllowEditScenarios ?? true; }
            set { bAllowEditScenarios = value; }
        }

        internal bool? bAllowFormatCells;
        /// <summary>
        /// Allow formatting of cells even if sheet is protected.
        /// </summary>
        public bool AllowFormatCells
        {
            get { return bAllowFormatCells ?? false; }
            set { bAllowFormatCells = value; }
        }

        internal bool? bAllowFormatColumns;
        /// <summary>
        /// Allow formatting of columns even if sheet is protected.
        /// </summary>
        public bool AllowFormatColumns
        {
            get { return bAllowFormatColumns ?? false; }
            set { bAllowFormatColumns = value; }
        }

        internal bool? bAllowFormatRows;
        /// <summary>
        /// Allow formatting of rows even if sheet is protected.
        /// </summary>
        public bool AllowFormatRows
        {
            get { return bAllowFormatRows ?? false; }
            set { bAllowFormatRows = value; }
        }

        internal bool? bAllowInsertColumns;
        /// <summary>
        /// Allow insertion of columns even if sheet is protected.
        /// </summary>
        public bool AllowInsertColumns
        {
            get { return bAllowInsertColumns ?? false; }
            set { bAllowInsertColumns = value; }
        }

        internal bool? bAllowInsertRows;
        /// <summary>
        /// Allow insertion of rows even if sheet is protected.
        /// </summary>
        public bool AllowInsertRows
        {
            get { return bAllowInsertRows ?? false; }
            set { bAllowInsertRows = value; }
        }

        internal bool? bAllowInsertHyperlinks;
        /// <summary>
        /// Allow insertion of hyperlinks even if sheet is protected.
        /// </summary>
        public bool AllowInsertHyperlinks
        {
            get { return bAllowInsertHyperlinks ?? false; }
            set { bAllowInsertHyperlinks = value; }
        }

        internal bool? bAllowDeleteColumns;
        /// <summary>
        /// Allow deletion of columns even if sheet is protected.
        /// </summary>
        public bool AllowDeleteColumns
        {
            get { return bAllowDeleteColumns ?? false; }
            set { bAllowDeleteColumns = value; }
        }

        internal bool? bAllowDeleteRows;
        /// <summary>
        /// Allow deletion of rows even if sheet is protected.
        /// </summary>
        public bool AllowDeleteRows
        {
            get { return bAllowDeleteRows ?? false; }
            set { bAllowDeleteRows = value; }
        }

        internal bool? bAllowSelectLockedCells;
        /// <summary>
        /// Allow selection of locked cells even if sheet is protected.
        /// </summary>
        public bool AllowSelectLockedCells
        {
            get { return bAllowSelectLockedCells ?? true; }
            set { bAllowSelectLockedCells = value; }
        }

        internal bool? bAllowSort;
        /// <summary>
        /// Allow sorting even if sheet is protected.
        /// </summary>
        public bool AllowSort
        {
            get { return bAllowSort ?? false; }
            set { bAllowSort = value; }
        }

        internal bool? bAllowAutoFilter;
        /// <summary>
        /// Allow use of autofilters even if sheet is protected.
        /// </summary>
        public bool AllowAutoFilter
        {
            get { return bAllowAutoFilter ?? false; }
            set { bAllowAutoFilter = value; }
        }

        internal bool? bAllowPivotTables;
        /// <summary>
        /// Allow use of pivot tables even if sheet is protected.
        /// </summary>
        public bool AllowPivotTables
        {
            get { return bAllowPivotTables ?? false; }
            set { bAllowPivotTables = value; }
        }

        internal bool? bAllowSelectUnlockedCells;
        /// <summary>
        /// Allow selection of unlocked cells even if sheet is protected.
        /// </summary>
        public bool AllowSelectUnlockedCells
        {
            get { return bAllowSelectUnlockedCells ?? true; }
            set { bAllowSelectUnlockedCells = value; }
        }

        /// <summary>
        /// Initializes an instance of SLSheetProtection.
        /// </summary>
        public SLSheetProtection()
        {
            this.SetAllNull();
        }

        internal void SetAllNull()
        {
            this.AlgorithmName = null;
            this.HashValue = null;
            this.SaltValue = null;
            this.SpinCount = null;
            this.Password = null;
            this.Sheet = null;
            this.bAllowEditObjects = null;
            this.bAllowEditScenarios = null;
            this.bAllowFormatCells = null;
            this.bAllowFormatColumns = null;
            this.bAllowFormatRows = null;
            this.bAllowInsertColumns = null;
            this.bAllowInsertRows = null;
            this.bAllowInsertHyperlinks = null;
            this.bAllowDeleteColumns = null;
            this.bAllowDeleteRows = null;
            this.bAllowSelectLockedCells = null;
            this.bAllowSort = null;
            this.bAllowAutoFilter = null;
            this.bAllowPivotTables = null;
            this.bAllowSelectUnlockedCells = null;
        }

        internal void FromSheetProtection(SheetProtection sp)
        {
            this.SetAllNull();
            if (sp.AlgorithmName != null) this.AlgorithmName = sp.AlgorithmName.Value;
            if (sp.HashValue != null) this.HashValue = sp.HashValue.Value;
            if (sp.SaltValue != null) this.SaltValue = sp.SaltValue.Value;
            if (sp.SpinCount != null) this.SpinCount = sp.SpinCount.Value;
            if (sp.Password != null) this.Password = sp.Password.Value;
            if (sp.Sheet != null) this.Sheet = sp.Sheet.Value;

            if (sp.Objects != null) this.AllowEditObjects = !sp.Objects.Value;
            if (sp.Scenarios != null) this.AllowEditScenarios = !sp.Scenarios.Value;
            if (sp.FormatCells != null) this.AllowFormatCells = !sp.FormatCells.Value;
            if (sp.FormatColumns != null) this.AllowFormatColumns = !sp.FormatColumns.Value;
            if (sp.FormatRows != null) this.AllowFormatRows = !sp.FormatRows.Value;
            if (sp.InsertColumns != null) this.AllowInsertColumns = !sp.InsertColumns.Value;
            if (sp.InsertRows != null) this.AllowInsertRows = !sp.InsertRows.Value;
            if (sp.InsertHyperlinks != null) this.AllowInsertHyperlinks = !sp.InsertHyperlinks.Value;
            if (sp.DeleteColumns != null) this.AllowDeleteColumns = !sp.DeleteColumns.Value;
            if (sp.DeleteRows != null) this.AllowDeleteRows = !sp.DeleteRows.Value;
            if (sp.SelectLockedCells != null) this.AllowSelectLockedCells = !sp.SelectLockedCells.Value;
            if (sp.Sort != null) this.AllowSort = !sp.Sort.Value;
            if (sp.AutoFilter != null) this.AllowAutoFilter = !sp.AutoFilter.Value;
            if (sp.PivotTables != null) this.AllowPivotTables = !sp.PivotTables.Value;
            if (sp.SelectUnlockedCells != null) this.AllowSelectUnlockedCells = !sp.SelectUnlockedCells.Value;
        }

        internal SheetProtection ToSheetProtection()
        {
            SheetProtection sp = new SheetProtection();
            if (this.AlgorithmName != null) sp.AlgorithmName = this.AlgorithmName;
            if (this.HashValue != null) sp.HashValue = this.HashValue;
            if (this.SaltValue != null) sp.SaltValue = this.SaltValue;
            if (this.SpinCount != null) sp.SpinCount = this.SpinCount.Value;
            if (this.Password != null) sp.Password = this.Password;
            if (this.Sheet != null && this.Sheet.Value != false) sp.Sheet = this.Sheet.Value;

            if (!this.AllowEditObjects != false) sp.Objects = !this.AllowEditObjects;
            if (!this.AllowEditScenarios != false) sp.Scenarios = !this.AllowEditScenarios;
            if (!this.AllowFormatCells != true) sp.FormatCells = !this.AllowFormatCells;
            if (!this.AllowFormatColumns != true) sp.FormatColumns = !this.AllowFormatColumns;
            if (!this.AllowFormatRows != true) sp.FormatRows = !this.AllowFormatRows;
            if (!this.AllowInsertColumns != true) sp.InsertColumns = !this.AllowInsertColumns;
            if (!this.AllowInsertRows != true) sp.InsertRows = !this.AllowInsertRows;
            if (!this.AllowInsertHyperlinks != true) sp.InsertHyperlinks = !this.AllowInsertHyperlinks;
            if (!this.AllowDeleteColumns != true) sp.DeleteColumns = !this.AllowDeleteColumns;
            if (!this.AllowDeleteRows != true) sp.DeleteRows = !this.AllowDeleteRows;
            if (!this.AllowSelectLockedCells != false) sp.SelectLockedCells = !this.AllowSelectLockedCells;
            if (!this.AllowSort != true) sp.Sort = !this.AllowSort;
            if (!this.AllowAutoFilter != true) sp.AutoFilter = !this.AllowAutoFilter;
            if (!this.AllowPivotTables != true) sp.PivotTables = !this.AllowPivotTables;
            if (!this.AllowSelectUnlockedCells != false) sp.SelectUnlockedCells = !this.AllowSelectUnlockedCells;

            return sp;
        }

        internal SLSheetProtection Clone()
        {
            SLSheetProtection sp = new SLSheetProtection();
            sp.AlgorithmName = this.AlgorithmName;
            sp.HashValue = this.HashValue;
            sp.SaltValue = this.SaltValue;
            sp.SpinCount = this.SpinCount;
            sp.Password = this.Password;
            sp.Sheet = this.Sheet;
            sp.bAllowEditObjects = this.bAllowEditObjects;
            sp.bAllowEditScenarios = this.bAllowEditScenarios;
            sp.bAllowFormatCells = this.bAllowFormatCells;
            sp.bAllowFormatColumns = this.bAllowFormatColumns;
            sp.bAllowFormatRows = this.bAllowFormatRows;
            sp.bAllowInsertColumns = this.bAllowInsertColumns;
            sp.bAllowInsertRows = this.bAllowInsertRows;
            sp.bAllowInsertHyperlinks = this.bAllowInsertHyperlinks;
            sp.bAllowDeleteColumns = this.bAllowDeleteColumns;
            sp.bAllowDeleteRows = this.bAllowDeleteRows;
            sp.bAllowSelectLockedCells = this.bAllowSelectLockedCells;
            sp.bAllowSort = this.bAllowSort;
            sp.bAllowAutoFilter = this.bAllowAutoFilter;
            sp.bAllowPivotTables = this.bAllowPivotTables;
            sp.bAllowSelectUnlockedCells = this.bAllowSelectUnlockedCells;

            return sp;
        }
    }
}
