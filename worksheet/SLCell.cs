using System;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// This is for information purposes only! This simulates the DocumentFormat.OpenXml.Spreadsheet.Cell class.
    /// </summary>
    public class SLCell
    {
        /// <summary>
        /// Indicates if the cell is truly empty. This is read-only.
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                return this.CellFormula == null
                    && (this.sCellText != null && this.sCellText.Length == 0)
                    && this.StyleIndex == 0 && this.DataType == CellValues.Number
                    && this.CellMetaIndex == 0 && this.ValueMetaIndex == 0
                    && !this.ShowPhonetic;
            }
        }

        //internal CellFormula Formula { get; set; }

        /// <summary>
        /// Cell formula.
        /// </summary>
        public SLCellFormula CellFormula { get; set; }

        private bool bToPreserveSpace;
        internal bool ToPreserveSpace
        {
            get { return bToPreserveSpace; }
        }

        private string sCellText;
        /// <summary>
        /// If this is null, the actual value is stored in NumericValue.
        /// </summary>
        public string CellText
        {
            get { return sCellText; }
            set
            {
                sCellText = value;
                bToPreserveSpace = SLTool.ToPreserveSpace(sCellText);

                if (value != null) fNumericValue = 0;
            }
        }

        /// <summary>
        /// Access this at your own peril! Only when CellText and NumericValue have to be set together! Probably! You've been warned!
        /// </summary>
        internal double fNumericValue;
        /// <summary>
        /// Use this value only when CellText is null.
        /// </summary>
        public double NumericValue
        {
            get { return fNumericValue; }
            set
            {
                fNumericValue = value;

                sCellText = null;
                bToPreserveSpace = false;
            }
        }

        // The logic will be to store boolean, numbers and shared string indices in
        // NumericValue. We'll actually use NumericValue if CellText is null.
        // This will keep memory low since Text will not always be used.
        // Most spreadsheets have numeric data. Consider "1234.56789". That's a 10
        // character string, but is always 8 bytes (?) if stored as a double.
        // In fact, any double is always stored as 8 bytes (thus the memory savings).
        // Plus it seems faster to assign and store a number to a double type than
        // storing the number in string form.

        // So. If CellText is null, it's a number type.
        // If CellText has some string in it, then we use that. And depending on the data type,
        // we'll interpret CellText differently.

        /// <summary>
        /// Style index.
        /// </summary>
        public uint StyleIndex { get; set; }

        /// <summary>
        /// Cell data type.
        /// </summary>
        public CellValues DataType { get; set; }

        /// <summary>
        /// Cell meta index.
        /// </summary>
        public uint CellMetaIndex { get; set; }

        /// <summary>
        /// Cell value meta index.
        /// </summary>
        public uint ValueMetaIndex { get; set; }

        /// <summary>
        /// Indicates if phonetic information should be shown.
        /// </summary>
        public bool ShowPhonetic { get; set; }

        internal SLCell()
        {
            this.SetAllNull();
        }

        internal void SetAllNull()
        {
            //this.Formula = null;
            this.CellFormula = null;

            this.bToPreserveSpace = false;
            this.sCellText = string.Empty;

            this.fNumericValue = 0;

            this.StyleIndex = 0;
            this.DataType = CellValues.Number;
            this.CellMetaIndex = 0;
            this.ValueMetaIndex = 0;
            this.ShowPhonetic = false;
        }

        internal void FromCell(Cell c)
        {
            this.SetAllNull();

            //if (c.CellFormula != null) this.Formula = (CellFormula)c.CellFormula.CloneNode(true);
            if (c.CellFormula != null)
            {
                this.CellFormula = new SLCellFormula();
                this.CellFormula.FromCellFormula(c.CellFormula);
            }

            if (c.StyleIndex != null) this.StyleIndex = c.StyleIndex.Value;

            if (c.DataType != null) this.DataType = c.DataType.Value;
            else this.DataType = CellValues.Number;

            if (c.CellValue != null) this.CellText = c.CellValue.Text ?? string.Empty;
            
            double fValue = 0;
            int iValue = 0;
            bool bValue = false;
            switch (this.DataType)
            {
                case CellValues.Number:
                    if (double.TryParse(this.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out fValue))
                    {
                        this.NumericValue = fValue;
                    }
                    break;
                case CellValues.SharedString:
                    if (int.TryParse(this.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out iValue))
                    {
                        this.NumericValue = iValue;
                    }
                    break;
                case CellValues.Boolean:
                    if (double.TryParse(this.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out fValue))
                    {
                        if (fValue > 0.5) this.NumericValue = 1;
                        else this.NumericValue = 0;
                    }
                    else if (bool.TryParse(this.CellText, out bValue))
                    {
                        if (bValue) this.NumericValue = 1;
                        else this.NumericValue = 0;
                    }
                    break;
            }
            
            if (c.CellMetaIndex != null) this.CellMetaIndex = c.CellMetaIndex.Value;
            if (c.ValueMetaIndex != null) this.ValueMetaIndex = c.ValueMetaIndex.Value;
            if (c.ShowPhonetic != null) this.ShowPhonetic = c.ShowPhonetic.Value;
        }

        internal Cell ToCell(string CellReference)
        {
            Cell c = new Cell();
            //if (this.Formula != null) c.CellFormula = this.Formula;
            if (this.CellFormula != null) c.CellFormula = this.CellFormula.ToCellFormula();

            if (this.CellText != null)
            {
                if (this.CellText.Length > 0)
                {
                    if (this.ToPreserveSpace)
                    {
                        c.CellValue = new CellValue(this.CellText)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        };
                    }
                    else
                    {
                        c.CellValue = new CellValue(this.CellText);
                    }
                }
            }
            else
            {
                // zero Text length
                if (this.DataType == CellValues.Number)
                {
                    c.CellValue = new CellValue(this.NumericValue.ToString(CultureInfo.InvariantCulture));
                }
                else if (this.DataType == CellValues.SharedString)
                {
                    c.CellValue = new CellValue(this.NumericValue.ToString("f0", CultureInfo.InvariantCulture));
                }
                else if (this.DataType == CellValues.Boolean)
                {
                    if (this.NumericValue > 0.5) c.CellValue = new CellValue("1");
                    else c.CellValue = new CellValue("0");
                }
            }

            c.CellReference = CellReference;
            if (this.StyleIndex > 0) c.StyleIndex = this.StyleIndex;
            if (this.DataType != CellValues.Number) c.DataType = this.DataType;
            if (this.CellMetaIndex > 0) c.CellMetaIndex = this.CellMetaIndex;
            if (this.ValueMetaIndex > 0) c.ValueMetaIndex = this.ValueMetaIndex;
            if (this.ShowPhonetic == true) c.ShowPhonetic = true;

            return c;
        }

        internal SLCell Clone()
        {
            SLCell cell = new SLCell();
            //if (this.Formula != null) cell.Formula = (CellFormula)this.Formula.CloneNode(true);
            if (this.CellFormula != null) cell.CellFormula = this.CellFormula.Clone();
            cell.bToPreserveSpace = this.bToPreserveSpace;
            cell.sCellText = this.sCellText;

            cell.fNumericValue = this.fNumericValue;

            cell.StyleIndex = this.StyleIndex;
            cell.DataType = this.DataType;
            cell.CellMetaIndex = this.CellMetaIndex;
            cell.ValueMetaIndex = this.ValueMetaIndex;
            cell.ShowPhonetic = this.ShowPhonetic;

            return cell;
        }
    }
}
