using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// <strong>Obsolete. </strong>Get the column name given the column index.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The column name.</returns>
        [Obsolete("Use SLConvert.ToColumnName() instead.")]
        public static string WhatIsColumnName(int ColumnIndex)
        {
            return SLTool.ToColumnName(ColumnIndex);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Get the column index given a cell reference or column name.
        /// </summary>
        /// <param name="Input">A cell reference such as "A1" or column name such as "A". If the input is invalid, then -1 is returned.</param>
        /// <returns>The column index.</returns>
        [Obsolete("Use SLConvert.ToColumnIndex() instead.")]
        public static int WhatIsColumnIndex(string Input)
        {
            return SLTool.ToColumnIndex(Input);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Get the cell reference given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The cell reference.</returns>
        [Obsolete("Use SLConvert.ToCellReference() instead.")]
        public static string WhatIsCellReference(int RowIndex, int ColumnIndex)
        {
            return SLTool.ToCellReference(RowIndex, ColumnIndex);
        }

        /// <summary>
        /// Get the row and column indices given a cell reference such as "C5". A return value indicates whether the conversion succeeded.
        /// </summary>
        /// <param name="CellReference">The cell reference in A1 format, such as "C5".</param>
        /// <param name="RowIndex">When this method returns, this contains the row index of the given cell reference if the conversion succeeded.</param>
        /// <param name="ColumnIndex">When this method returns, this contains the column index of the given cell reference if the conversion succeeded.</param>
        /// <returns>True if the conversion succeeded. False otherwise.</returns>
        public static bool WhatIsRowColumnIndex(string CellReference, out int RowIndex, out int ColumnIndex)
        {
            RowIndex = -1;
            ColumnIndex = -1;
            return SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out RowIndex, out ColumnIndex);
        }

        /// <summary>
        /// Get the width of the specified column.
        /// </summary>
        /// <param name="Unit">The unit of the width to be returned.</param>
        /// <param name="ColumnName">The column name, such as "A".</param>
        /// <returns>The width in the specified unit type.</returns>
        public double GetWidth(SLMeasureUnitTypeValues Unit, string ColumnName)
        {
            int iColumnIndex = -1;
            iColumnIndex = SLTool.ToColumnIndex(ColumnName);

            return this.GetWidth(Unit, iColumnIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the width of the specified columns.
        /// </summary>
        /// <param name="Unit">The unit of the width to be returned.</param>
        /// <param name="StartColumnName">The column name of the start column.</param>
        /// <param name="EndColumnName">The column name of the end column.</param>
        /// <returns>The width in the specified unit type.</returns>
        public double GetWidth(SLMeasureUnitTypeValues Unit, string StartColumnName, string EndColumnName)
        {
            int iStartColumnIndex = -1;
            int iEndColumnIndex = -1;
            iStartColumnIndex = SLTool.ToColumnIndex(StartColumnName);
            iEndColumnIndex = SLTool.ToColumnIndex(EndColumnName);

            return this.GetWidth(Unit, iStartColumnIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Get the width of the specified column.
        /// </summary>
        /// <param name="Unit">The unit of the width to be returned.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The width in the specified unit type.</returns>
        public double GetWidth(SLMeasureUnitTypeValues Unit, int ColumnIndex)
        {
            return this.GetWidth(Unit, ColumnIndex, ColumnIndex);
        }

        /// <summary>
        /// Get the width of the specified columns.
        /// </summary>
        /// <param name="Unit">The unit of the width to be returned.</param>
        /// <param name="StartColumnIndex">The column index of the start column.</param>
        /// <param name="EndColumnIndex">The column index of the end column.</param>
        /// <returns>The width in the specified unit type.</returns>
        public double GetWidth(SLMeasureUnitTypeValues Unit, int StartColumnIndex, int EndColumnIndex)
        {
            if (StartColumnIndex < 1) StartColumnIndex = 1;
            if (StartColumnIndex > SLConstants.ColumnLimit) StartColumnIndex = SLConstants.ColumnLimit;
            if (EndColumnIndex < 1) EndColumnIndex = 1;
            if (EndColumnIndex > SLConstants.ColumnLimit) EndColumnIndex = SLConstants.ColumnLimit;

            long lWidth = 0;
            int i = 0;
            SLColumnProperties cp;
            for (i = StartColumnIndex; i <= EndColumnIndex; ++i)
            {
                if (slws.ColumnProperties.ContainsKey(i))
                {
                    cp = slws.ColumnProperties[i];
                    lWidth += cp.WidthInEMU;
                }
                else
                {
                    lWidth += slws.SheetFormatProperties.DefaultColumnWidthInEMU;
                }
            }

            double result = 0;
            switch (Unit)
            {
                case SLMeasureUnitTypeValues.Centimeter:
                    result = SLConvert.FromEmuToCentimeter((double)lWidth);
                    break;
                case SLMeasureUnitTypeValues.Emu:
                    result = (double)lWidth;
                    break;
                case SLMeasureUnitTypeValues.Inch:
                    result = SLConvert.FromEmuToInch((double)lWidth);
                    break;
                case SLMeasureUnitTypeValues.Point:
                    result = SLConvert.FromEmuToPoint((double)lWidth);
                    break;
            }

            return result;
        }

        /// <summary>
        /// Get the height of the specified row.
        /// </summary>
        /// <param name="Unit">The unit of the height to be returned.</param>
        /// <param name="RowIndex">The row index.</param>
        /// <returns>The height in the specified unit type.</returns>
        public double GetHeight(SLMeasureUnitTypeValues Unit, int RowIndex)
        {
            return this.GetHeight(Unit, RowIndex, RowIndex);
        }

        /// <summary>
        /// Get the height of specified rows.
        /// </summary>
        /// <param name="Unit">The unit of the height to be returned.</param>
        /// <param name="StartRowIndex">The row index of the start row.</param>
        /// <param name="EndRowIndex">The row index of the end row.</param>
        /// <returns>The height in the specified unit type.</returns>
        public double GetHeight(SLMeasureUnitTypeValues Unit, int StartRowIndex, int EndRowIndex)
        {
            if (StartRowIndex < 1) StartRowIndex = 1;
            if (StartRowIndex > SLConstants.RowLimit) StartRowIndex = SLConstants.RowLimit;
            if (EndRowIndex < 1) EndRowIndex = 1;
            if (EndRowIndex > SLConstants.RowLimit) EndRowIndex = SLConstants.RowLimit;

            long lHeight = 0;
            int i = 0;
            SLRowProperties rp;
            for (i = StartRowIndex; i <= EndRowIndex; ++i)
            {
                if (slws.RowProperties.ContainsKey(i))
                {
                    rp = slws.RowProperties[i];
                    lHeight += rp.HeightInEMU;
                }
                else
                {
                    lHeight += slws.SheetFormatProperties.DefaultRowHeightInEMU;
                }
            }

            double result = 0;
            switch (Unit)
            {
                case SLMeasureUnitTypeValues.Centimeter:
                    result = SLConvert.FromEmuToCentimeter((double)lHeight);
                    break;
                case SLMeasureUnitTypeValues.Emu:
                    result = (double)lHeight;
                    break;
                case SLMeasureUnitTypeValues.Inch:
                    result = SLConvert.FromEmuToInch((double)lHeight);
                    break;
                case SLMeasureUnitTypeValues.Point:
                    result = SLConvert.FromEmuToPoint((double)lHeight);
                    break;
            }

            return result;
        }

        /// <summary>
        /// Indicates if there's an existing defined name given a name.
        /// </summary>
        /// <param name="Name">Name of defined name to check.</param>
        /// <returns>True if the defined name exists. False otherwise.</returns>
        public bool HasDefinedName(string Name)
        {
            bool result = false;
            if (wbp.Workbook.DefinedNames != null)
            {
                foreach (var child in wbp.Workbook.DefinedNames.Elements<DefinedName>())
                {
                    if (child.Name != null && child.Name.Value.Equals(Name, StringComparison.OrdinalIgnoreCase))
                    {
                        result = true;
                        break;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Set a given defined name. If it doesn't exist, a new defined name is created. If it exists, then the existing defined name is overwritten.
        /// </summary>
        /// <param name="Name">Name of defined name. Note that it cannot be a valid cell reference such as A1. It also cannot start with "_xlnm" because it's reserved.</param>
        /// <param name="Text">The reference/content text of the defined name. For example, Sheet1!$A$1:$C$3</param>
        /// <returns>True if the given defined name is created or an existing defined name is overwritten. False otherwise.</returns>
        public bool SetDefinedName(string Name, string Text)
        {
            return SetDefinedName(Name, Text, string.Empty, string.Empty);
        }

        /// <summary>
        /// Set a given defined name. If it doesn't exist, a new defined name is created. If it exists, then the existing defined name is overwritten.
        /// </summary>
        /// <param name="Name">Name of defined name. Note that it cannot be a valid cell reference such as A1. It also cannot start with "_xlnm" because it's reserved.</param>
        /// <param name="Text">The reference/content text of the defined name. For example, Sheet1!$A$1:$C$3</param>
        /// <param name="Comment">Comment for the defined name.</param>
        /// <returns>True if the given defined name is created or an existing defined name is overwritten. False otherwise.</returns>
        public bool SetDefinedName(string Name, string Text, string Comment)
        {
            return SetDefinedName(Name, Text, Comment, string.Empty);
        }

        /// <summary>
        /// Set a given defined name. If it doesn't exist, a new defined name is created. If it exists, then the existing defined name is overwritten.
        /// </summary>
        /// <param name="Name">Name of defined name. Note that it cannot be a valid cell reference such as A1. It also cannot start with "_xlnm" because it's reserved.</param>
        /// <param name="Text">The reference/content text of the defined name. For example, Sheet1!$A$1:$C$3</param>
        /// <param name="Comment">Comment for the defined name.</param>
        /// <param name="Scope">The name of the worksheet that the defined name is effective in.</param>
        /// <returns>True if the given defined name is created or an existing defined name is overwritten. False otherwise.</returns>
        public bool SetDefinedName(string Name, string Text, string Comment, string Scope)
        {
            Name = Name.Trim();
            if (SLTool.IsCellReference(Name))
            {
                return false;
            }

            // these are reserved names
            if (Name.StartsWith("_xlnm")) return false;

            if (Text.StartsWith("="))
            {
                if (Text.Length > 1) Text = Text.Substring(1);
                else Text = "\"=\"";
            }

            uint? iLocalSheetId = null;
            for (int i = 0; i < slwb.Sheets.Count; ++i)
            {
                if (slwb.Sheets[i].Name.Equals(Scope, StringComparison.OrdinalIgnoreCase))
                {
                    iLocalSheetId = (uint)i;
                    break;
                }
            }

            bool bFound = false;
            SLDefinedName dn = new SLDefinedName(Name);
            dn.Text = Text;
            if (Comment != null && Comment.Length > 0) dn.Comment = Comment;
            if (iLocalSheetId != null) dn.LocalSheetId = iLocalSheetId.Value;
            foreach (SLDefinedName d in slwb.DefinedNames)
            {
                if (d.Name.Equals(Name, StringComparison.OrdinalIgnoreCase))
                {
                    bFound = true;
                    d.Text = Text;
                    if (Comment != null && Comment.Length > 0) d.Comment = Comment;
                    break;
                }
            }

            if (!bFound)
            {
                slwb.DefinedNames.Add(dn);
            }

            return true;
        }

        /// <summary>
        /// Get reference/content text of existing defined name.
        /// </summary>
        /// <param name="Name">Name of existing defined name.</param>
        /// <returns>Reference/content text of defined name. An empty string is returned if the given defined name doesn't exist.</returns>
        public string GetDefinedNameText(string Name)
        {
            string result = string.Empty;
            foreach (SLDefinedName d in slwb.DefinedNames)
            {
                if (d.Name.Equals(Name, StringComparison.OrdinalIgnoreCase))
                {
                    result = d.Text;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Get the comment of existing defined name.
        /// </summary>
        /// <param name="Name">Name of existing defined name.</param>
        /// <returns>The comment of the defined name. An empty string is returned if the given defined name doesn't exist, or there's no comment.</returns>
        public string GetDefinedNameComment(string Name)
        {
            string result = string.Empty;
            foreach (SLDefinedName d in slwb.DefinedNames)
            {
                if (d.Name.Equals(Name, StringComparison.OrdinalIgnoreCase))
                {
                    result = d.Comment ?? string.Empty;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Delete a defined name if it exists.
        /// </summary>
        /// <param name="Name">Name of defined name.</param>
        /// <returns>True if specified name is deleted. False otherwise.</returns>
        public bool DeleteDefinedName(string Name)
        {
            bool result = false;
            for (int i = 0; i < slwb.DefinedNames.Count; ++i)
            {
                if (slwb.DefinedNames[i].Name.Equals(Name, StringComparison.OrdinalIgnoreCase))
                {
                    result = true;
                    slwb.DefinedNames.RemoveAt(i);
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Get a list of existing defined names.
        /// </summary>
        /// <returns>A list of defined names.</returns>
        public List<SLDefinedName> GetDefinedNames()
        {
            return GetDefinedNames(false);
        }

        /// <summary>
        /// Get a list of existing defined names, filtered by whether the defined name is a reserved name or not.
        /// </summary>
        /// <param name="IncludeReserved">True to include reserved names. False otherwise. A reserved name starts with "_xlnm".</param>
        /// <returns>A list of defined names.</returns>
        public List<SLDefinedName> GetDefinedNames(bool IncludeReserved)
        {
            List<SLDefinedName> result = new List<SLDefinedName>();

            if (IncludeReserved)
            {
                foreach (SLDefinedName dn in slwb.DefinedNames)
                {
                    result.Add(dn.Clone());
                }
            }
            else
            {
                foreach (SLDefinedName dn in slwb.DefinedNames)
                {
                    if (!dn.Name.StartsWith("_xlnm")) result.Add(dn.Clone());
                }
            }
            
            return result;
        }

        /// <summary>
        /// Set the print area on the currently selected worksheet given a corner cell of the print area and the opposite corner cell.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        public void SetPrintArea(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex);
            SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex);

            this.SetAddPrintArea(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, false);
        }

        /// <summary>
        /// Set the print area on the currently selected worksheet given a corner cell of the print area and the opposite corner cell.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        public void SetPrintArea(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            this.SetAddPrintArea(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, false);
        }

        /// <summary>
        /// Adds a print area to the existing print area on the currently selected worksheet given a corner cell of the print area and the opposite corner cell.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        public void AddToPrintArea(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex);
            SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex);

            this.SetAddPrintArea(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, true);
        }

        /// <summary>
        /// Adds a print area to the existing print area on the currently selected worksheet given a corner cell of the print area and the opposite corner cell.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        public void AddToPrintArea(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            this.SetAddPrintArea(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, true);
        }

        private void SetAddPrintArea(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool ToAdd)
        {
            if (StartRowIndex < 1) StartRowIndex = 1;
            if (StartRowIndex > SLConstants.RowLimit) StartRowIndex = SLConstants.RowLimit;
            if (StartColumnIndex < 1) StartColumnIndex = 1;
            if (StartColumnIndex > SLConstants.ColumnLimit) StartColumnIndex = SLConstants.ColumnLimit;
            if (EndRowIndex < 1) EndRowIndex = 1;
            if (EndRowIndex > SLConstants.RowLimit) EndRowIndex = SLConstants.RowLimit;
            if (EndColumnIndex < 1) EndColumnIndex = 1;
            if (EndColumnIndex > SLConstants.ColumnLimit) EndColumnIndex = SLConstants.ColumnLimit;

            // no overlapping checked.

            int iSheetPosition = 0;
            int i;
            for (i = 0; i < slwb.Sheets.Count; ++i)
            {
                if (slwb.Sheets[i].Name.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    iSheetPosition = i;
                    break;
                }
            }

            string sPrintArea = string.Empty;
            if (StartRowIndex == EndRowIndex && StartColumnIndex == EndColumnIndex)
            {
                // why would you print just one cell? Even Excel questions this with a message box...
                sPrintArea = SLTool.ToCellReference(gsSelectedWorksheetName, StartRowIndex, StartColumnIndex, true);
            }
            else
            {
                sPrintArea = SLTool.ToCellRange(gsSelectedWorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, true);
            }

            bool bFound = false;
            for (i = 0; i < slwb.DefinedNames.Count; ++i)
            {
                if (slwb.DefinedNames[i].Name.Equals(SLConstants.PrintAreaDefinedName, StringComparison.OrdinalIgnoreCase)
                    && slwb.DefinedNames[i].LocalSheetId != null
                    && slwb.DefinedNames[i].LocalSheetId.Value == iSheetPosition)
                {
                    bFound = true;
                    if (ToAdd)
                    {
                        slwb.DefinedNames[i].Text = string.Format("{0},{1}", slwb.DefinedNames[i].Text, sPrintArea);
                    }
                    else
                    {
                        slwb.DefinedNames[i].Text = sPrintArea;
                    }
                }
            }

            if (!bFound)
            {
                SLDefinedName dn = new SLDefinedName(SLConstants.PrintAreaDefinedName);
                dn.LocalSheetId = (uint)iSheetPosition;
                dn.Text = sPrintArea;
                slwb.DefinedNames.Add(dn);
            }
        }

        /// <summary>
        /// Clears existing print areas on the currently selected worksheet.
        /// </summary>
        public void ClearPrintArea()
        {
            int iSheetPosition = 0;
            int i;
            for (i = 0; i < slwb.Sheets.Count; ++i)
            {
                if (slwb.Sheets[i].Name.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    iSheetPosition = i;
                    break;
                }
            }

            for (i = slwb.DefinedNames.Count - 1; i >= 0; --i)
            {
                if (slwb.DefinedNames[i].Name.Equals(SLConstants.PrintAreaDefinedName, StringComparison.OrdinalIgnoreCase)
                    && slwb.DefinedNames[i].LocalSheetId != null
                    && slwb.DefinedNames[i].LocalSheetId.Value == iSheetPosition)
                {
                    slwb.DefinedNames.RemoveAt(i);
                    break;
                }
            }
        }

        /// <summary>
        /// Insert a hyperlink.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="HyperlinkType">The type of hyperlink.</param>
        /// <param name="Address">The URL for web pages, the file path for existing files, a cell reference (such as Sheet1!A1 or Sheet1!A1:B5), a defined name or an email address. NOTE: Do NOT include the "mailto:" portion for email addresses.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertHyperlink(string CellReference, SLHyperlinkTypeValues HyperlinkType, string Address)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return InsertHyperlink(iRowIndex, iColumnIndex, HyperlinkType, Address, null, null, false);
        }

        /// <summary>
        /// Insert a hyperlink.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="HyperlinkType">The type of hyperlink.</param>
        /// <param name="Address">The URL for web pages, the file path for existing files, a cell reference (such as Sheet1!A1 or Sheet1!A1:B5), a defined name or an email address. NOTE: Do NOT include the "mailto:" portion for email addresses.</param>
        /// <param name="OverwriteExistingCell">True to overwrite the existing cell value with the hyperlink display text. False otherwise.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertHyperlink(string CellReference, SLHyperlinkTypeValues HyperlinkType, string Address, bool OverwriteExistingCell)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return InsertHyperlink(iRowIndex, iColumnIndex, HyperlinkType, Address, null, null, OverwriteExistingCell);
        }

        /// <summary>
        /// Insert a hyperlink.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="HyperlinkType">The type of hyperlink.</param>
        /// <param name="Address">The URL for web pages, the file path for existing files, a cell reference (such as Sheet1!A1 or Sheet1!A1:B5), a defined name or an email address. NOTE: Do NOT include the "mailto:" portion for email addresses.</param>
        /// <param name="Display">The display text. Set null or an empty string to use the default.</param>
        /// <param name="ToolTip">The tooltip (or screentip) text. Set null or an empty string to ignore this.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertHyperlink(string CellReference, SLHyperlinkTypeValues HyperlinkType, string Address, string Display, string ToolTip)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return InsertHyperlink(iRowIndex, iColumnIndex, HyperlinkType, Address, Display, ToolTip, false);
        }

        /// <summary>
        /// Insert a hyperlink.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="HyperlinkType">The type of hyperlink.</param>
        /// <param name="Address">The URL for web pages, the file path for existing files, a cell reference (such as Sheet1!A1 or Sheet1!A1:B5), a defined name or an email address. NOTE: Do NOT include the "mailto:" portion for email addresses.</param>
        /// <param name="Display">The display text. Set null or an empty string to use the default.</param>
        /// <param name="ToolTip">The tooltip (or screentip) text. Set null or an empty string to ignore this.</param>
        /// <param name="OverwriteExistingCell">True to overwrite the existing cell value with the hyperlink display text. False otherwise.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertHyperlink(string CellReference, SLHyperlinkTypeValues HyperlinkType, string Address, string Display, string ToolTip, bool OverwriteExistingCell)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return InsertHyperlink(iRowIndex, iColumnIndex, HyperlinkType, Address, Display, ToolTip, OverwriteExistingCell);
        }

        /// <summary>
        /// Insert a hyperlink.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="HyperlinkType">The type of hyperlink.</param>
        /// <param name="Address">The URL for web pages, the file path for existing files, a cell reference (such as Sheet1!A1 or Sheet1!A1:B5), a defined name or an email address. NOTE: Do NOT include the "mailto:" portion for email addresses.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertHyperlink(int RowIndex, int ColumnIndex, SLHyperlinkTypeValues HyperlinkType, string Address)
        {
            return InsertHyperlink(RowIndex, ColumnIndex, HyperlinkType, Address, null, null, false);
        }

        /// <summary>
        /// Insert a hyperlink.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="HyperlinkType">The type of hyperlink.</param>
        /// <param name="Address">The URL for web pages, the file path for existing files, a cell reference (such as Sheet1!A1 or Sheet1!A1:B5), a defined name or an email address. NOTE: Do NOT include the "mailto:" portion for email addresses.</param>
        /// <param name="OverwriteExistingCell">True to overwrite the existing cell value with the hyperlink display text. False otherwise.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertHyperlink(int RowIndex, int ColumnIndex, SLHyperlinkTypeValues HyperlinkType, string Address, bool OverwriteExistingCell)
        {
            return InsertHyperlink(RowIndex, ColumnIndex, HyperlinkType, Address, null, null, OverwriteExistingCell);
        }

        /// <summary>
        /// Insert a hyperlink.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="HyperlinkType">The type of hyperlink.</param>
        /// <param name="Address">The URL for web pages, the file path for existing files, a cell reference (such as Sheet1!A1 or Sheet1!A1:B5), a defined name or an email address. NOTE: Do NOT include the "mailto:" portion for email addresses.</param>
        /// <param name="Display">The display text. Set null or an empty string to use the default.</param>
        /// <param name="ToolTip">The tooltip (or screentip) text. Set null or an empty string to ignore this.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertHyperlink(int RowIndex, int ColumnIndex, SLHyperlinkTypeValues HyperlinkType, string Address, string Display, string ToolTip)
        {
            return InsertHyperlink(RowIndex, ColumnIndex, HyperlinkType, Address, Display, ToolTip, false);
        }

        // TODO: Hyperlink cell range

        /// <summary>
        /// Insert a hyperlink.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="HyperlinkType">The type of hyperlink.</param>
        /// <param name="Address">The URL for web pages, the file path for existing files, a cell reference (such as Sheet1!A1 or Sheet1!A1:B5), a defined name or an email address. NOTE: Do NOT include the "mailto:" portion for email addresses.</param>
        /// <param name="Display">The display text. Set null or an empty string to use the default.</param>
        /// <param name="ToolTip">The tooltip (or screentip) text. Set null or an empty string to ignore this.</param>
        /// <param name="OverwriteExistingCell">True to overwrite the existing cell value with the hyperlink display text. False otherwise.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertHyperlink(int RowIndex, int ColumnIndex, SLHyperlinkTypeValues HyperlinkType, string Address, string Display, string ToolTip, bool OverwriteExistingCell)
        {
            if (RowIndex < 1 || RowIndex > SLConstants.RowLimit) return false;
            if (ColumnIndex < 1 || ColumnIndex > SLConstants.ColumnLimit) return false;

            SLHyperlink hl = new SLHyperlink();
            hl.IsNew = true;

            hl.Reference = new SLCellPointRange(RowIndex, ColumnIndex, RowIndex, ColumnIndex);

            switch (HyperlinkType)
            {
                case SLHyperlinkTypeValues.EmailAddress:
                    hl.IsExternal = true;
                    hl.HyperlinkUri = string.Format("mailto:{0}", Address);
                    hl.HyperlinkUriKind = UriKind.Absolute;
                    break;
                case SLHyperlinkTypeValues.FilePath:
                    hl.IsExternal = true;
                    hl.HyperlinkUri = Address;
                    // assume if it starts with ../ or ./ it's a relative path.
                    hl.HyperlinkUriKind = Address.StartsWith(".") ? UriKind.Relative : UriKind.Absolute;
                    break;
                case SLHyperlinkTypeValues.InternalDocumentLink:
                    hl.IsExternal = false;
                    hl.Location = Address;
                    break;
                case SLHyperlinkTypeValues.Url:
                    hl.IsExternal = true;
                    hl.HyperlinkUri = Address;
                    hl.HyperlinkUriKind = UriKind.Absolute;
                    break;
            }

            if (Display == null)
            {
                hl.Display = Address;
            }
            else
            {
                if (Display.Length == 0)
                {
                    hl.Display = Address;
                }
                else
                {
                    hl.Display = Display;
                }
            }

            if (ToolTip != null && ToolTip.Length > 0) hl.ToolTip = ToolTip;

            SLCell c;
            SLStyle style;
            if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
            {
                c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                style = new SLStyle();
                if (c.StyleIndex < listStyle.Count) style.FromHash(listStyle[(int)c.StyleIndex]);
                else style.FromHash(listStyle[0]);
                style.SetFontUnderline(UnderlineValues.Single);
                style.SetFontColor(SLThemeColorIndexValues.Hyperlink);
                c.StyleIndex = (uint)this.SaveToStylesheet(style.ToHash());

                if (OverwriteExistingCell)
                {
                    // in case there's a formula
                    c.CellFormula = null;
                    c.DataType = CellValues.SharedString;
                    c.CellText = this.DirectSaveToSharedStringTable(hl.Display).ToString(CultureInfo.InvariantCulture);
                }
                // else don't have to do anything

                slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
            }
            else
            {
                c = new SLCell();

                style = new SLStyle();
                style.FromHash(listStyle[0]);
                style.SetFontUnderline(UnderlineValues.Single);
                style.SetFontColor(SLThemeColorIndexValues.Hyperlink);
                c.StyleIndex = (uint)this.SaveToStylesheet(style.ToHash());

                c.DataType = CellValues.SharedString;
                c.CellText = this.DirectSaveToSharedStringTable(hl.Display).ToString(CultureInfo.InvariantCulture);
                slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
            }

            slws.Hyperlinks.Add(hl);

            return true;
        }

        /// <summary>
        /// Remove an existing hyperlink.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        public void RemoveHyperlink(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex);

            this.RemoveHyperlink(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Remove an existing hyperlink.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        public void RemoveHyperlink(int RowIndex, int ColumnIndex)
        {
            if (RowIndex < 1 || RowIndex > SLConstants.RowLimit) return;
            if (ColumnIndex < 1 || ColumnIndex > SLConstants.ColumnLimit) return;

            // I'm assuming hyperlinks are tied to just one cell. Apparently,
            // you can assign a block of cells as the hyperlink.
            // Excel removes the cells of the hyperlink that are empty. I'm not going to even try...

            List<SLCellPointRange> listdelete = new List<SLCellPointRange>();

            int i, j;
            WorksheetPart wsp;
            string sRelId;
            HyperlinkRelationship hlrel;
            if (!IsNewWorksheet)
            {
                for (i = slws.Hyperlinks.Count - 1; i >= 0; --i)
                {
                    if (slws.Hyperlinks[i].Reference.StartRowIndex <= RowIndex
                        && RowIndex <= slws.Hyperlinks[i].Reference.EndRowIndex
                        && slws.Hyperlinks[i].Reference.StartColumnIndex <= ColumnIndex
                        && ColumnIndex <= slws.Hyperlinks[i].Reference.EndColumnIndex)
                    {
                        if (slws.Hyperlinks[i].IsNew)
                        {
                            slws.Hyperlinks.RemoveAt(i);
                        }
                        else
                        {
                            if (slws.Hyperlinks[i].IsExternal
                                && slws.Hyperlinks[i].Id != null
                                && slws.Hyperlinks[i].Id.Length > 0)
                            {
                                sRelId = slws.Hyperlinks[i].Id;
                                if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
                                {
                                    wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                                    hlrel = wsp.HyperlinkRelationships.Where(hlid => hlid.Id == sRelId).FirstOrDefault();
                                    if (hlrel != null)
                                    {
                                        wsp.DeleteReferenceRelationship(hlrel);
                                    }
                                }
                            }

                            slws.Hyperlinks.RemoveAt(i);
                        }

                        listdelete.Add(new SLCellPointRange(
                            slws.Hyperlinks[i].Reference.StartRowIndex,
                            slws.Hyperlinks[i].Reference.StartColumnIndex,
                            slws.Hyperlinks[i].Reference.EndRowIndex,
                            slws.Hyperlinks[i].Reference.EndColumnIndex));
                    }
                }
            }
            else
            {
                // if it's a new worksheet, all hyperlinks are new.
                // Most importantly, no hyperlink relationships were added, so
                // we can just remove from the SLWorksheet list.
                // Start from the end because we'll be deleting.
                for (i = slws.Hyperlinks.Count - 1; i >= 0; --i)
                {
                    if (slws.Hyperlinks[i].Reference.StartRowIndex <= RowIndex
                        && RowIndex <= slws.Hyperlinks[i].Reference.EndRowIndex
                        && slws.Hyperlinks[i].Reference.StartColumnIndex <= ColumnIndex
                        && ColumnIndex <= slws.Hyperlinks[i].Reference.EndColumnIndex)
                    {
                        slws.Hyperlinks.RemoveAt(i);

                        listdelete.Add(new SLCellPointRange(
                            slws.Hyperlinks[i].Reference.StartRowIndex,
                            slws.Hyperlinks[i].Reference.StartColumnIndex,
                            slws.Hyperlinks[i].Reference.EndRowIndex,
                            slws.Hyperlinks[i].Reference.EndColumnIndex));
                    }
                }
            }

            if (listdelete.Count > 0)
            {
                // remove hyperlink style
                SLCell c;
                foreach (SLCellPointRange ptrange in listdelete)
                {
                    for (i = ptrange.StartRowIndex; i <= ptrange.EndRowIndex; ++i)
                    {
                        for (j = ptrange.StartColumnIndex; j <= ptrange.EndColumnIndex; ++j)
                        {
                            if (slws.CellWarehouse.Exists(i, j))
                            {
                                c = slws.CellWarehouse.Cells[i][j];
                                c.StyleIndex = 0;
                                slws.CellWarehouse.SetValue(i, j, c);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Refresh pivot table data on load or open of spreadsheet. This applies to all pivot tables on the currently selected worksheet.
        /// </summary>
        public void RefreshPivotTableOnLoad()
        {
            // Note that this refreshes the pivot table cache. The actual data of the pivot
            // table in the worksheet isn't refreshed. Meaning if you get cell data from a cell
            // in the pivot table, you'll get the old value (at least that's what I believe happens).
            // This tells Excel to refresh the data, and the new data is displayed. When you then save,
            // Excel saves the newly refreshed data.
            // The cache is stored separately from the worksheet data.
            if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
            {
                WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                if (wsp.PivotTableParts != null)
                {
                    foreach (PivotTablePart ptp in wsp.PivotTableParts)
                    {
                        if (ptp.PivotTableCacheDefinitionPart != null
                            && ptp.PivotTableCacheDefinitionPart.PivotCacheDefinition != null)
                        {
                            ptp.PivotTableCacheDefinitionPart.PivotCacheDefinition.RefreshOnLoad = true;
                        }
                    }
                }
            }
        }
    }
}
