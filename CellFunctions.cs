using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// This cleans up all the SLCell objects with default values.
        /// This can happen if an SLCell was assigned with say a style but no value.
        /// Then subsequently, the style is removed (set to default), thus the cell is empty.
        /// </summary>
        internal void CleanUpReallyEmptyCells()
        {
            // Realistically speaking, there shouldn't be a lot of cells with default values.
            // But we don't want SheetData to be cluttered, and also this saves maybe a few bytes.

            List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
            List<int> listColumnKeys;

            foreach (int rowkey in listRowKeys)
            {
                listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                foreach (int colkey in listColumnKeys)
                {
                    if (slws.CellWarehouse.Cells[rowkey][colkey].IsEmpty)
                    {
                        slws.CellWarehouse.Remove(rowkey, colkey);
                    }
                }

                if (slws.CellWarehouse.Cells[rowkey].Count == 0)
                {
                    slws.CellWarehouse.Cells.Remove(rowkey);
                }
            }
        }

        internal void CheckAndClearSharedCellFormulaIfNeedTo(int RowIndex, int ColumnIndex)
        {
            if (this.slws.SharedCellFormulas.Count > 0)
            {
                bool bFound = false;
                foreach (SLSharedCellFormula scf in this.slws.SharedCellFormulas.Values)
                {
                    if (RowIndex == scf.BaseCellRowIndex && ColumnIndex == scf.BaseCellColumnIndex)
                    {
                        bFound = true;
                        break;
                    }
                }

                if (bFound)
                {
                    this.FlattenAllSharedCellFormula();
                }
            }
        }

        /// <summary>
        /// Get existing cells in the currently selected worksheet. WARNING: This is only a snapshot. Any changes made to the returned result are not used.
        /// </summary>
        /// <returns>A Dictionary of existing cells.</returns>
        public Dictionary<int, Dictionary<int, SLCell>> GetCells()
        {
            Dictionary<int, Dictionary<int, SLCell>> result = new Dictionary<int, Dictionary<int, SLCell>>();

            List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
            List<int> listColumnKeys;

            foreach (int rowkey in listRowKeys)
            {
                result.Add(rowkey, new Dictionary<int, SLCell>());
                listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                foreach (int colkey in listColumnKeys)
                {
                    result[rowkey].Add(colkey, slws.CellWarehouse.Cells[rowkey][colkey].Clone());
                }
            }

            return result;
        }

        /// <summary>
        /// Indicates if the cell value exists.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>True if it exists. False otherwise.</returns>
        public bool HasCellValue(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return HasCellValue(iRowIndex, iColumnIndex, false);
        }

        /// <summary>
        /// Indicates if the cell value exists.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>True if it exists. False otherwise.</returns>
        public bool HasCellValue(int RowIndex, int ColumnIndex)
        {
            return HasCellValue(RowIndex, ColumnIndex, false);
        }

        /// <summary>
        /// Indicates if the cell value exists.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="IncludeCellFormula">True if having a cell formula counts as well. False otherwise.</param>
        /// <returns>True if it exists. False otherwise.</returns>
        public bool HasCellValue(string CellReference, bool IncludeCellFormula)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return HasCellValue(iRowIndex, iColumnIndex, IncludeCellFormula);
        }

        /// <summary>
        /// Indicates if the cell value exists.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="IncludeCellFormula">True if having a cell formula counts as well. False otherwise.</param>
        /// <returns>True if it exists. False otherwise (also if out of bounds).</returns>
        public bool HasCellValue(int RowIndex, int ColumnIndex, bool IncludeCellFormula)
        {
            if (RowIndex < 1 || RowIndex > SLConstants.RowLimit) return false;
            if (ColumnIndex < 1 || ColumnIndex > SLConstants.ColumnLimit) return false;

            bool result = false;
            if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
            {
                SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                if (c.CellText == null)
                {
                    // if it's null, then it's using the numeric value portion, hence non-empty.
                    result = true;
                }
                else
                {
                    // else not null but we check for empty string
                    if (c.CellText.Length > 0) result = true;
                }

                if (IncludeCellFormula)
                {
                    result |= (c.CellFormula != null);
                }
            }

            return result;
        }

        /// <summary>
        /// Indicates if the cell has an error. WARNING: SpreadsheetLight does not have a formula calculation engine, so only existing errors are reported.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>True if there's a cell error. False otherwise (also if out of bounds).</returns>
        public bool HasCellError(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return HasCellError(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Indicates if the cell has an error. WARNING: SpreadsheetLight does not have a formula calculation engine, so only existing errors are reported.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>True if there's a cell error. False otherwise (also if out of bounds).</returns>
        public bool HasCellError(int RowIndex, int ColumnIndex)
        {
            if (RowIndex < 1 || RowIndex > SLConstants.RowLimit) return false;
            if (ColumnIndex < 1 || ColumnIndex > SLConstants.ColumnLimit) return false;

            bool result = false;
            if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
            {
                SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];

                // assume if the data type is Error, it's probably an error
                // but we also just check all the possible error values just in case
                if (c.DataType == CellValues.Error
                    || (c.CellText != null && string.Equals(c.CellText, SLConstants.ErrorDivisionByZero, StringComparison.OrdinalIgnoreCase))
                    || (c.CellText != null && string.Equals(c.CellText, SLConstants.ErrorNA, StringComparison.OrdinalIgnoreCase))
                    || (c.CellText != null && string.Equals(c.CellText, SLConstants.ErrorName, StringComparison.OrdinalIgnoreCase))
                    || (c.CellText != null && string.Equals(c.CellText, SLConstants.ErrorNull, StringComparison.OrdinalIgnoreCase))
                    || (c.CellText != null && string.Equals(c.CellText, SLConstants.ErrorNumber, StringComparison.OrdinalIgnoreCase))
                    || (c.CellText != null && string.Equals(c.CellText, SLConstants.ErrorReference, StringComparison.OrdinalIgnoreCase))
                    || (c.CellText != null && string.Equals(c.CellText, SLConstants.ErrorValue, StringComparison.OrdinalIgnoreCase))
                    )
                {
                    result = true;
                }
            }

            return result;
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, bool Data)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return SetCellValue(iRowIndex, iColumnIndex, Data);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, bool Data)
        {
            if (!SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                return false;
            }

            this.CheckAndClearSharedCellFormulaIfNeedTo(RowIndex, ColumnIndex);

            SLCell c;
            if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
            {
                c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
            }
            else
            {
                c = new SLCell();
                if (slws.RowProperties.ContainsKey(RowIndex))
                {
                    c.StyleIndex = slws.RowProperties[RowIndex].StyleIndex;
                }
                else if (slws.ColumnProperties.ContainsKey(ColumnIndex))
                {
                    c.StyleIndex = slws.ColumnProperties[ColumnIndex].StyleIndex;
                }
            }
            c.DataType = CellValues.Boolean;
            c.NumericValue = Data ? 1 : 0;
            slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);

            return true;
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, float Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, float Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, double Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Data, null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, double Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Data, null);
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, decimal Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, decimal Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, byte Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, byte Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, short Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, short Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, ushort Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, ushort Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, int Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, int Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, uint Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, uint Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, long Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, long Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, ulong Data)
        {
            return SetCellValueNumberFinal(CellReference, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. If you plan to store a percentage value, set this as the value divided by 100. For example, to store 2.78%, set this value as 0.0278. Remember to set the cell style format code (say "0.00%")!</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, ulong Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, true, Convert.ToDouble(Data), null);
        }

        /// <summary>
        /// Set the cell value given a cell reference and a numeric value in string form. Use this when the source data is numeric and is already in string form and parsing the data into numeric form is undesirable. Note that the numeric string must be in invariant-culture mode, so "123456.789" is the accepted form even if the current culture displays that as "123456,789".
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValueNumeric(string CellReference, string Data)
        {
            return SetCellValueNumberFinal(CellReference, false, 0, Data);
        }

        /// <summary>
        /// Set the cell value given the row index and column index and a numeric value in string form. Use this when the source data is numeric and is already in string form and parsing the data into numeric form is undesirable. Note that the numeric string must be in invariant-culture mode, so "123456.789" is the accepted form even if the current culture displays that as "123456,789".
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValueNumeric(int RowIndex, int ColumnIndex, string Data)
        {
            return SetCellValueNumberFinal(RowIndex, ColumnIndex, false, 0, Data);
        }

        /// <summary>
        /// Set the cell value given a cell reference. Be sure to follow up with a date format style.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, DateTime Data)
        {
            return SetCellValue(CellReference, Data, string.Empty, false);
        }

        /// <summary>
        /// Set the cell value given a cell reference. Be sure to follow up with a date format style.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data.</param>
        /// <param name="For1904Epoch">True if using 1 Jan 1904 as the date epoch. False if using 1 Jan 1900 as the date epoch. This is independent of the workbook's Date1904 property.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, DateTime Data, bool For1904Epoch)
        {
            return SetCellValue(CellReference, Data, string.Empty, For1904Epoch);
        }

        /// <summary>
        /// Set the cell value given a cell reference. Be sure to follow up with a date format style.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data.</param>
        /// <param name="Format">The format string used if the given date is before the date epoch. A date before the date epoch is stored as a string, so the date precision is only as good as the format string. For example, "dd/MM/yyyy HH:mm:ss" is more precise than "dd/MM/yyyy" because the latter loses information about the hours, minutes and seconds.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, DateTime Data, string Format)
        {
            return SetCellValue(CellReference, Data, Format, false);
        }

        /// <summary>
        /// Set the cell value given a cell reference. Be sure to follow up with a date format style.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data.</param>
        /// <param name="Format">The format string used if the given date is before the date epoch. A date before the date epoch is stored as a string, so the date precision is only as good as the format string. For example, "dd/MM/yyyy HH:mm:ss" is more precise than "dd/MM/yyyy" because the latter loses information about the hours, minutes and seconds.</param>
        /// <param name="For1904Epoch">True if using 1 Jan 1904 as the date epoch. False if using 1 Jan 1900 as the date epoch. This is independent of the workbook's Date1904 property.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, DateTime Data, string Format, bool For1904Epoch)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return SetCellValue(iRowIndex, iColumnIndex, Data, Format, For1904Epoch);
        }

        /// <summary>
        /// Set the cell value given the row index and column index. Be sure to follow up with a date format style.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, DateTime Data)
        {
            return SetCellValue(RowIndex, ColumnIndex, Data, string.Empty, false);
        }

        /// <summary>
        /// Set the cell value given the row index and column index. Be sure to follow up with a date format style.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data.</param>
        /// <param name="For1904Epoch">True if using 1 Jan 1904 as the date epoch. False if using 1 Jan 1900 as the date epoch. This is independent of the workbook's Date1904 property.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, DateTime Data, bool For1904Epoch)
        {
            return SetCellValue(RowIndex, ColumnIndex, Data, string.Empty, For1904Epoch);
        }

        /// <summary>
        /// Set the cell value given the row index and column index. Be sure to follow up with a date format style.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data.</param>
        /// <param name="Format">The format string used if the given date is before the date epoch. A date before the date epoch is stored as a string, so the date precision is only as good as the format string. For example, "dd/MM/yyyy HH:mm:ss" is more precise than "dd/MM/yyyy" because the latter loses information about the hours, minutes and seconds.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, DateTime Data, string Format)
        {
            return SetCellValue(RowIndex, ColumnIndex, Data, Format, false);
        }

        /// <summary>
        /// Set the cell value given the row index and column index. Be sure to follow up with a date format style.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data.</param>
        /// <param name="Format">The format string used if the given date is before the date epoch. A date before the date epoch is stored as a string, so the date precision is only as good as the format string. For example, "dd/MM/yyyy HH:mm:ss" is more precise than "dd/MM/yyyy" because the latter loses information about the hours, minutes and seconds.</param>
        /// <param name="For1904Epoch">True if using 1 Jan 1904 as the date epoch. False if using 1 Jan 1900 as the date epoch. This is independent of the workbook's Date1904 property.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, DateTime Data, string Format, bool For1904Epoch)
        {
            if (!SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                return false;
            }

            this.CheckAndClearSharedCellFormulaIfNeedTo(RowIndex, ColumnIndex);

            SLCell c;
            if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
            {
                c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
            }
            else
            {
                c = new SLCell();
                if (slws.RowProperties.ContainsKey(RowIndex))
                {
                    c.StyleIndex = slws.RowProperties[RowIndex].StyleIndex;
                }
                else if (slws.ColumnProperties.ContainsKey(ColumnIndex))
                {
                    c.StyleIndex = slws.ColumnProperties[ColumnIndex].StyleIndex;
                }
            }

            if (For1904Epoch) slwb.WorkbookProperties.Date1904 = true;

            double fDateTime = SLTool.CalculateDaysFromEpoch(Data, For1904Epoch);
            // see CalculateDaysFromEpoch to see why there's a difference
            double fDateCheck = For1904Epoch ? 0.0 : 1.0;

            if (fDateTime < fDateCheck)
            {
                // given datetime is earlier than epoch
                // So we set date to string format
                c.DataType = CellValues.SharedString;
                c.NumericValue = this.DirectSaveToSharedStringTable(Data.ToString(Format));
                slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
            }
            else
            {
                c.DataType = CellValues.Number;
                c.NumericValue = fDateTime;
                slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
            }

            return true;
        }

        private bool SetCellValueNumberFinal(string CellReference, bool IsNumeric, double NumericValue, string NumberData)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return SetCellValueNumberFinal(iRowIndex, iColumnIndex, IsNumeric, NumericValue, NumberData);
        }

        private bool SetCellValueNumberFinal(int RowIndex, int ColumnIndex, bool IsNumeric, double NumericValue, string NumberData)
        {
            if (!SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                return false;
            }

            this.CheckAndClearSharedCellFormulaIfNeedTo(RowIndex, ColumnIndex);

            SLCell c;
            if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
            {
                c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
            }
            else
            {
                c = new SLCell();
                if (slws.RowProperties.ContainsKey(RowIndex))
                {
                    c.StyleIndex = slws.RowProperties[RowIndex].StyleIndex;
                }
                else if (slws.ColumnProperties.ContainsKey(ColumnIndex))
                {
                    c.StyleIndex = slws.ColumnProperties[ColumnIndex].StyleIndex;
                }
            }
            c.DataType = CellValues.Number;
            if (IsNumeric) c.NumericValue = NumericValue;
            else c.CellText = NumberData;
            slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);

            return true;
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data in rich text.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, SLRstType Data)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return SetCellValue(iRowIndex, iColumnIndex, Data.ToInlineString());
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data in rich text.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, SLRstType Data)
        {
            return SetCellValue(RowIndex, ColumnIndex, Data.ToInlineString());
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data. Try the SLRstType class for easy InlineString generation.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, InlineString Data)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return SetCellValue(iRowIndex, iColumnIndex, Data);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data. Try the SLRstType class for easy InlineString generation.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, InlineString Data)
        {
            if (!SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                return false;
            }

            this.CheckAndClearSharedCellFormulaIfNeedTo(RowIndex, ColumnIndex);

            SLCell c;
            if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
            {
                c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
            }
            else
            {
                c = new SLCell();
                if (slws.RowProperties.ContainsKey(RowIndex))
                {
                    c.StyleIndex = slws.RowProperties[RowIndex].StyleIndex;
                }
                else if (slws.ColumnProperties.ContainsKey(ColumnIndex))
                {
                    c.StyleIndex = slws.ColumnProperties[ColumnIndex].StyleIndex;
                }
            }
            c.DataType = CellValues.SharedString;
            c.NumericValue = this.DirectSaveToSharedStringTable(Data);
            slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);

            return true;
        }

        /// <summary>
        /// Set the cell value given a cell reference.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if the cell reference is invalid. True otherwise.</returns>
        public bool SetCellValue(string CellReference, string Data)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return SetCellValue(iRowIndex, iColumnIndex, Data);
        }

        /// <summary>
        /// Set the cell value given the row index and column index.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Data">The cell value data.</param>
        /// <returns>False if either the row index or column index (or both) are invalid. True otherwise.</returns>
        public bool SetCellValue(int RowIndex, int ColumnIndex, string Data)
        {
            if (!SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                return false;
            }

            this.CheckAndClearSharedCellFormulaIfNeedTo(RowIndex, ColumnIndex);

            SLCell c;
            if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
            {
                c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
            }
            else
            {
                // if there's no existing cell, then we don't have to assign
                // a new cell when the data string is empty
                if (string.IsNullOrEmpty(Data)) return true;

                c = new SLCell();
                if (slws.RowProperties.ContainsKey(RowIndex))
                {
                    c.StyleIndex = slws.RowProperties[RowIndex].StyleIndex;
                }
                else if (slws.ColumnProperties.ContainsKey(ColumnIndex))
                {
                    c.StyleIndex = slws.ColumnProperties[ColumnIndex].StyleIndex;
                }
            }

            if (string.IsNullOrEmpty(Data))
            {
                c.DataType = CellValues.Number;
                c.CellText = string.Empty;
                slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
            }
            else if (Data.StartsWith("="))
            {
                // in case it's just one equal sign
                if (Data.Equals("=", StringComparison.OrdinalIgnoreCase))
                {
                    c.DataType = CellValues.SharedString;
                    c.NumericValue = this.DirectSaveToSharedStringTable("=");
                    slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
                }
                else
                {
                    // For simplicity, we're gonna assume that if it starts with an equal sign, it's a formula.

                    // TODO Formula calculation engine. Actually nope...
                    c.DataType = CellValues.Number;
                    //c.Formula = new CellFormula(slxe.Write(Data.Substring(1)));
                    c.CellFormula = new SLCellFormula();
                    //c.CellFormula.FormulaText = SLTool.XmlWrite(Data.Substring(1));
                    // apparently, you don't need to XML-escape double quotes otherwise there's an error.
                    c.CellFormula.FormulaText = Data.Substring(1);
                    c.CellText = string.Empty;
                    slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
                }
            }
            else if (Data.StartsWith("'"))
            {
                c.DataType = CellValues.SharedString;
                c.NumericValue = this.DirectSaveToSharedStringTable(SLTool.XmlWrite(Data.Substring(1), gbThrowExceptionsIfAny));
                slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
            }
            else
            {
                c.DataType = CellValues.SharedString;
                c.NumericValue = this.DirectSaveToSharedStringTable(SLTool.XmlWrite(Data, gbThrowExceptionsIfAny));
                slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
            }

            return true;
        }

        /// <summary>
        /// Get the cell value as a boolean. If the cell value wasn't originally a boolean value, the return value is undetermined (but is by default false).
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>A boolean cell value.</returns>
        public bool GetCellValueAsBoolean(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return GetCellValueAsBoolean(iRowIndex, iColumnIndex, false);
        }

        /// <summary>
        /// Get the cell value as a boolean. If the cell value wasn't originally a boolean value, the return value is undetermined (but is by default false).
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="TryForceParse">Set true to force any cell value that looks like a boolean to be returned as a boolean. This means text stored as "1" or "TRUE" will also be considered a boolean true. Set to false to only consider true (haha pun!) booleans. The default is false.</param>
        /// <returns>A boolean cell value.</returns>
        public bool GetCellValueAsBoolean(string CellReference, bool TryForceParse)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return GetCellValueAsBoolean(iRowIndex, iColumnIndex, TryForceParse);
        }

        /// <summary>
        /// Get the cell value as a boolean. If the cell value wasn't originally a boolean value, the return value is undetermined (but is by default false).
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>A boolean cell value.</returns>
        public bool GetCellValueAsBoolean(int RowIndex, int ColumnIndex)
        {
            return GetCellValueAsBoolean(RowIndex, ColumnIndex, false);
        }

        /// <summary>
        /// Get the cell value as a boolean. If the cell value wasn't originally a boolean value, the return value is undetermined (but is by default false).
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="TryForceParse">Set true to force any cell value that looks like a boolean to be returned as a boolean. This means text stored as "1" or "TRUE" will also be considered a boolean true. Set to false to only consider true (haha pun!) booleans. The default is false.</param>
        /// <returns>A boolean cell value.</returns>
        public bool GetCellValueAsBoolean(int RowIndex, int ColumnIndex, bool TryForceParse)
        {
            bool result = false;

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.DataType == CellValues.Boolean)
                    {
                        double fValue = 0;
                        if (c.CellText != null)
                        {
                            if (double.TryParse(c.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out fValue))
                            {
                                if (fValue > 0.5) result = true;
                                else result = false;
                            }
                            else
                            {
                                bool.TryParse(c.CellText, out result);
                            }
                        }
                        else
                        {
                            if (c.NumericValue > 0.5) result = true;
                            else result = false;
                        }
                    }
                    else if (TryForceParse && c.CellText != null)
                    {
                        string sText = string.Empty;
                        if (c.DataType == CellValues.String)
                        {
                            sText = SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny);
                        }
                        else if (c.DataType == CellValues.SharedString)
                        {
                            SLRstType rst = new SLRstType(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                            int index;
                            try
                            {
                                index = int.Parse(c.CellText);
                                if (index >= 0 && index < listSharedString.Count)
                                {
                                    rst.FromHash(listSharedString[index]);
                                    sText = rst.ToPlainString();
                                }
                                else
                                {
                                    sText = SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny);
                                }
                            }
                            catch (Exception e)
                            {
                                if (gbThrowExceptionsIfAny)
                                {
                                    throw e;
                                }
                                else
                                {
                                    // something terrible just happened. We'll just use whatever's in the cell...
                                    sText = SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny);
                                }
                            }
                        }

                        if (sText.Length > 0)
                        {
                            if (sText.Equals("1") || sText.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
                            {
                                result = true;
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get the cell value as a 32-bit integer. If the cell value wasn't originally an integer, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>A 32-bit integer cell value.</returns>
        public Int32 GetCellValueAsInt32(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return 0;
            }

            return GetCellValueAsInt32(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell value as a 32-bit integer. If the cell value wasn't originally an integer, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>A 32-bit integer cell value.</returns>
        public Int32 GetCellValueAsInt32(int RowIndex, int ColumnIndex)
        {
            Int32 result = 0;

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.DataType == CellValues.Number)
                    {
                        if (c.CellText != null)
                        {
                            Int32.TryParse(c.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                        }
                        else
                        {
                            result = Convert.ToInt32(c.NumericValue);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get the cell value as an unsigned 32-bit integer. If the cell value wasn't originally an integer, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>An unsigned 32-bit integer cell value.</returns>
        public UInt32 GetCellValueAsUInt32(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return 0;
            }

            return GetCellValueAsUInt32(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell value as an unsigned 32-bit integer. If the cell value wasn't originally an integer, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>An unsigned 32-bit integer cell value.</returns>
        public UInt32 GetCellValueAsUInt32(int RowIndex, int ColumnIndex)
        {
            UInt32 result = 0;

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.DataType == CellValues.Number)
                    {
                        if (c.CellText != null)
                        {
                            UInt32.TryParse(c.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                        }
                        else
                        {
                            result = Convert.ToUInt32(c.NumericValue);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get the cell value as a 64-bit integer. If the cell value wasn't originally an integer, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>A 64-bit integer cell value.</returns>
        public Int64 GetCellValueAsInt64(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return 0;
            }

            return GetCellValueAsInt64(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell value as a 64-bit integer. If the cell value wasn't originally an integer, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>A 64-bit integer cell value.</returns>
        public Int64 GetCellValueAsInt64(int RowIndex, int ColumnIndex)
        {
            Int64 result = 0;

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.DataType == CellValues.Number)
                    {
                        if (c.CellText != null)
                        {
                            Int64.TryParse(c.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                        }
                        else
                        {
                            result = Convert.ToInt64(c.NumericValue);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get the cell value as an unsigned 64-bit integer. If the cell value wasn't originally an integer, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>An unsigned 64-bit integer cell value.</returns>
        public UInt64 GetCellValueAsUInt64(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return 0;
            }

            return GetCellValueAsUInt64(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell value as an unsigned 64-bit integer. If the cell value wasn't originally an integer, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>An unsigned 64-bit integer cell value.</returns>
        public UInt64 GetCellValueAsUInt64(int RowIndex, int ColumnIndex)
        {
            UInt64 result = 0;

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.DataType == CellValues.Number)
                    {
                        if (c.CellText != null)
                        {
                            UInt64.TryParse(c.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                        }
                        else
                        {
                            result = Convert.ToUInt64(c.NumericValue);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get the cell value as a double precision floating point number. If the cell value wasn't originally a floating point number, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>A double precision floating point number cell value.</returns>
        public double GetCellValueAsDouble(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return 0.0;
            }

            return GetCellValueAsDouble(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell value as a double precision floating point number. If the cell value wasn't originally a floating point number, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>A double precision floating point number cell value.</returns>
        public double GetCellValueAsDouble(int RowIndex, int ColumnIndex)
        {
            double result = 0.0;

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.DataType == CellValues.Number)
                    {
                        if (c.CellText != null)
                        {
                            double.TryParse(c.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                        }
                        else
                        {
                            result = c.NumericValue;
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get the cell value as a System.Decimal value. If the cell value wasn't originally an integer or floating point number, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>A System.Decimal cell value.</returns>
        public decimal GetCellValueAsDecimal(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return 0m;
            }

            return GetCellValueAsDecimal(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell value as a System.Decimal value. If the cell value wasn't originally an integer or floating point number, the return value is undetermined (but is by default 0).
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>A System.Decimal cell value.</returns>
        public decimal GetCellValueAsDecimal(int RowIndex, int ColumnIndex)
        {
            decimal result = 0m;

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.DataType == CellValues.Number)
                    {
                        if (c.CellText != null)
                        {
                            decimal.TryParse(c.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                        }
                        else
                        {
                            result = Convert.ToDecimal(c.NumericValue);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get the cell value as a System.DateTime value. If the cell value wasn't originally a date/time value, the return value is undetermined.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>A System.DateTime cell value.</returns>
        public DateTime GetCellValueAsDateTime(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                if (slwb.WorkbookProperties.Date1904) return SLConstants.Epoch1904();
                else return SLConstants.Epoch1900();
            }

            return GetCellValueAsDateTime(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell value as a System.DateTime value. If the cell value wasn't originally a date/time value, the return value is undetermined.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>A System.DateTime cell value.</returns>
        public DateTime GetCellValueAsDateTime(int RowIndex, int ColumnIndex)
        {
            return GetCellValueAsDateTime(RowIndex, ColumnIndex, string.Empty, false);
        }

        /// <summary>
        /// Get the cell value as a System.DateTime value. If the cell value wasn't originally a date/time value, the return value is undetermined.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="For1904Epoch">True if using 1 Jan 1904 as the date epoch. False if using 1 Jan 1900 as the date epoch. This is independent of the workbook's Date1904 property.</param>
        /// <returns>A System.DateTime cell value.</returns>
        public DateTime GetCellValueAsDateTime(string CellReference, bool For1904Epoch)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                if (slwb.WorkbookProperties.Date1904) return SLConstants.Epoch1904();
                else return SLConstants.Epoch1900();
            }

            return GetCellValueAsDateTime(iRowIndex, iColumnIndex, For1904Epoch);
        }

        /// <summary>
        /// Get the cell value as a System.DateTime value. If the cell value wasn't originally a date/time value, the return value is undetermined.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="For1904Epoch">True if using 1 Jan 1904 as the date epoch. False if using 1 Jan 1900 as the date epoch. This is independent of the workbook's Date1904 property.</param>
        /// <returns>A System.DateTime cell value.</returns>
        public DateTime GetCellValueAsDateTime(int RowIndex, int ColumnIndex, bool For1904Epoch)
        {
            return GetCellValueAsDateTime(RowIndex, ColumnIndex, string.Empty, For1904Epoch);
        }

        /// <summary>
        /// Get the cell value as a System.DateTime value. If the cell value wasn't originally a date/time value, the return value is undetermined.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Format">The format string used to parse the date value in the cell if the date is before the date epoch. A date before the date epoch is stored as a string, so the date precision is only as good as the format string. For example, "dd/MM/yyyy HH:mm:ss" is more precise than "dd/MM/yyyy" because the latter loses information about the hours, minutes and seconds.</param>
        /// <returns>A System.DateTime cell value.</returns>
        public DateTime GetCellValueAsDateTime(string CellReference, string Format)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                if (slwb.WorkbookProperties.Date1904) return SLConstants.Epoch1904();
                else return SLConstants.Epoch1900();
            }

            return GetCellValueAsDateTime(iRowIndex, iColumnIndex, Format);
        }

        /// <summary>
        /// Get the cell value as a System.DateTime value. If the cell value wasn't originally a date/time value, the return value is undetermined.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Format">The format string used to parse the date value in the cell if the date is before the date epoch. A date before the date epoch is stored as a string, so the date precision is only as good as the format string. For example, "dd/MM/yyyy HH:mm:ss" is more precise than "dd/MM/yyyy" because the latter loses information about the hours, minutes and seconds.</param>
        /// <returns>A System.DateTime cell value.</returns>
        public DateTime GetCellValueAsDateTime(int RowIndex, int ColumnIndex, string Format)
        {
            return GetCellValueAsDateTime(RowIndex, ColumnIndex, Format, false);
        }

        /// <summary>
        /// Get the cell value as a System.DateTime value. If the cell value wasn't originally a date/time value, the return value is undetermined.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <param name="Format">The format string used to parse the date value in the cell if the date is before the date epoch. A date before the date epoch is stored as a string, so the date precision is only as good as the format string. For example, "dd/MM/yyyy HH:mm:ss" is more precise than "dd/MM/yyyy" because the latter loses information about the hours, minutes and seconds.</param>
        /// <param name="For1904Epoch">True if using 1 Jan 1904 as the date epoch. False if using 1 Jan 1900 as the date epoch. This is independent of the workbook's Date1904 property.</param>
        /// <returns>A System.DateTime cell value.</returns>
        public DateTime GetCellValueAsDateTime(string CellReference, string Format, bool For1904Epoch)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                if (slwb.WorkbookProperties.Date1904) return SLConstants.Epoch1904();
                else return SLConstants.Epoch1900();
            }

            return GetCellValueAsDateTime(iRowIndex, iColumnIndex, Format, For1904Epoch);
        }

        /// <summary>
        /// Get the cell value as a System.DateTime value. If the cell value wasn't originally a date/time value, the return value is undetermined.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="Format">The format string used to parse the date value in the cell if the date is before the date epoch. A date before the date epoch is stored as a string, so the date precision is only as good as the format string. For example, "dd/MM/yyyy HH:mm:ss" is more precise than "dd/MM/yyyy" because the latter loses information about the hours, minutes and seconds.</param>
        /// <param name="For1904Epoch">True if using 1 Jan 1904 as the date epoch. False if using 1 Jan 1900 as the date epoch. This is independent of the workbook's Date1904 property.</param>
        /// <returns>A System.DateTime cell value.</returns>
        public DateTime GetCellValueAsDateTime(int RowIndex, int ColumnIndex, string Format, bool For1904Epoch)
        {
            DateTime dt;
            if (For1904Epoch) dt = SLConstants.Epoch1904();
            else dt = SLConstants.Epoch1900();

            // If the cell data type is Number, then it's on or after the epoch.
            // If it's a string or a shared string, then a string representation of the date
            // is stored, where the date is before the epoch. Then we parse the string to
            // get the date.

            double fDateOffset = 0.0;
            string sDate = string.Empty;

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.DataType == CellValues.Number)
                    {
                        if (c.CellText != null)
                        {
                            if (double.TryParse(c.CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out fDateOffset))
                            {
                                dt = SLTool.CalculateDateTimeFromDaysFromEpoch(fDateOffset, For1904Epoch);
                            }
                        }
                        else
                        {
                            dt = SLTool.CalculateDateTimeFromDaysFromEpoch(c.NumericValue, For1904Epoch);
                        }
                    }
                    else if (c.DataType == CellValues.SharedString)
                    {
                        SLRstType rst = new SLRstType(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                        int index = 0;
                        try
                        {
                            if (c.CellText != null)
                            {
                                index = int.Parse(c.CellText);
                            }
                            else
                            {
                                index = Convert.ToInt32(c.NumericValue);
                            }
                            
                            if (index >= 0 && index < listSharedString.Count)
                            {
                                rst.FromHash(listSharedString[index]);
                                sDate = rst.ToPlainString();

                                if (Format.Length > 0)
                                {
                                    dt = DateTime.ParseExact(sDate, Format, CultureInfo.InvariantCulture);
                                }
                                else
                                {
                                    dt = DateTime.Parse(sDate, CultureInfo.InvariantCulture);
                                }
                            }
                            // no else part, because there's nothing we can do!
                            // Just return the default date value...
                        }
                        catch (Exception e)
                        {
                            if (gbThrowExceptionsIfAny) throw e;
                            // else something terrible just happened. (the shared string index probably
                            // isn't even correct!) Don't do anything...
                        }
                    }
                    else if (c.DataType == CellValues.String)
                    {
                        sDate = c.CellText ?? string.Empty;
                        try
                        {
                            if (Format.Length > 0)
                            {
                                dt = DateTime.ParseExact(sDate, Format, CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                dt = DateTime.Parse(sDate, CultureInfo.InvariantCulture);
                            }
                        }
                        catch (Exception e)
                        {
                            if (gbThrowExceptionsIfAny) throw e;
                            // else don't need to do anything. Just return the default date value.
                        }
                    }
                }
            }

            return dt;
        }

        /// <summary>
        /// Get the cell value as a string.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>A string cell value.</returns>
        public string GetCellValueAsString(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return string.Empty;
            }

            return GetCellValueAsString(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell value as a string.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>A string cell value.</returns>
        public string GetCellValueAsString(int RowIndex, int ColumnIndex)
        {
            string result = string.Empty;
            int index = 0;
            SLRstType rst = new SLRstType(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.CellText != null)
                    {
                        if (c.DataType == CellValues.String)
                        {
                            result = SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny);
                        }
                        else if (c.DataType == CellValues.SharedString)
                        {
                            try
                            {
                                index = int.Parse(c.CellText);
                                if (index >= 0 && index < listSharedString.Count)
                                {
                                    rst.FromHash(listSharedString[index]);
                                    result = rst.ToPlainString();
                                }
                                else
                                {
                                    result = SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny);
                                }
                            }
                            catch (Exception e)
                            {
                                if (gbThrowExceptionsIfAny)
                                {
                                    throw e;
                                }
                                else
                                {
                                    // something terrible just happened. We'll just use whatever's in the cell...
                                    result = SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny);
                                }
                            }
                        }
                        else if (c.DataType == CellValues.InlineString)
                        {
                            // there shouldn't be any inline strings
                            // because they'd already be transferred to shared strings
                            // but just in case...
                        }
                        else
                        {
                            result = SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny);
                        }
                    }
                    else
                    {
                        if (c.DataType == CellValues.Number)
                        {
                            result = c.NumericValue.ToString(CultureInfo.InvariantCulture);
                        }
                        else if (c.DataType == CellValues.SharedString)
                        {
                            index = Convert.ToInt32(c.NumericValue);
                            if (index >= 0 && index < listSharedString.Count)
                            {
                                rst.FromHash(listSharedString[index]);
                                result = rst.ToPlainString();
                            }
                            else
                            {
                                result = SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny);
                            }
                        }
                        else if (c.DataType == CellValues.Boolean)
                        {
                            if (c.NumericValue > 0.5) result = "TRUE";
                            else result = "FALSE";
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Get the cell value as a rich text string (SLRstType).
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>An SLRstType cell value.</returns>
        public SLRstType GetCellValueAsRstType(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return new SLRstType(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            }

            return GetCellValueAsRstType(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell value as a rich text string (SLRstType).
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>An SLRstType cell value.</returns>
        public SLRstType GetCellValueAsRstType(int RowIndex, int ColumnIndex)
        {
            SLRstType rst = new SLRstType(SimpleTheme.MajorLatinFont, SimpleTheme.MinorLatinFont, SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
            int index = 0;

            if (SLTool.CheckRowColumnIndexLimit(RowIndex, ColumnIndex))
            {
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (c.CellText != null)
                    {
                        if (c.DataType == CellValues.String)
                        {
                            rst.SetText(SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny));
                        }
                        else if (c.DataType == CellValues.SharedString)
                        {
                            try
                            {
                                index = int.Parse(c.CellText);
                                if (index >= 0 && index < listSharedString.Count)
                                {
                                    rst.FromHash(listSharedString[index]);
                                }
                                else
                                {
                                    rst.SetText(SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny));
                                }
                            }
                            catch (Exception e)
                            {
                                if (gbThrowExceptionsIfAny)
                                {
                                    throw e;
                                }
                                else
                                {
                                    // something terrible just happened. We'll just use whatever's in the cell...
                                    rst.SetText(SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny));
                                }
                            }
                        }
                        //else if (c.DataType == CellValues.InlineString)
                        //{
                        //    // there shouldn't be any inline strings
                        //    // because they'd already be transferred to shared strings
                        //}
                        else
                        {
                            rst.SetText(SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny));
                        }
                    }
                    else
                    {
                        if (c.DataType == CellValues.Number)
                        {
                            rst.SetText(c.NumericValue.ToString(CultureInfo.InvariantCulture));
                        }
                        else if (c.DataType == CellValues.SharedString)
                        {
                            index = Convert.ToInt32(c.NumericValue);
                            if (index >= 0 && index < listSharedString.Count)
                            {
                                rst.FromHash(listSharedString[index]);
                            }
                            else
                            {
                                rst.SetText(SLTool.XmlRead(c.CellText, gbThrowExceptionsIfAny));
                            }
                        }
                        else if (c.DataType == CellValues.Boolean)
                        {
                            if (c.NumericValue > 0.5) rst.SetText("TRUE");
                            else rst.SetText("FALSE");
                        }
                    }
                }
            }

            return rst.Clone();
        }

        /// <summary>
        /// Set the active cell for the currently selected worksheet.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetActiveCell(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return false;
            }

            return this.SetActiveCell(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Set the active cell for the currently selected worksheet.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetActiveCell(int RowIndex, int ColumnIndex)
        {
            if (RowIndex < 1 || RowIndex > SLConstants.RowLimit || ColumnIndex < 1 || ColumnIndex > SLConstants.ColumnLimit)
                return false;

            slws.ActiveCell = new SLCellPoint(RowIndex, ColumnIndex);

            int i, j;
            SLSheetView sv;
            SLSelection sel;
            if (slws.SheetViews.Count == 0)
            {
                // if it's A1, I'm not going to do anything. It's the default!
                if (RowIndex != 1 || ColumnIndex != 1)
                {
                    sv = new SLSheetView();
                    sv.WorkbookViewId = 0;
                    sel = new SLSelection();
                    sel.ActiveCell = SLTool.ToCellReference(RowIndex, ColumnIndex);
                    sel.SequenceOfReferences.Add(new SLCellPointRange(RowIndex, ColumnIndex, RowIndex, ColumnIndex));
                    sv.Selections.Add(sel);

                    slws.SheetViews.Add(sv);
                }
            }
            else
            {
                bool bFound = false;
                PaneValues vActivePane = PaneValues.TopLeft;
                for (i = 0; i < slws.SheetViews.Count; ++i)
                {
                    if (slws.SheetViews[i].WorkbookViewId == 0)
                    {
                        bFound = true;
                        if (slws.SheetViews[i].Selections.Count == 0)
                        {
                            if (RowIndex != 1 || ColumnIndex != 1)
                            {
                                sel = new SLSelection();
                                sel.ActiveCell = SLTool.ToCellReference(RowIndex, ColumnIndex);
                                sel.SequenceOfReferences.Add(new SLCellPointRange(RowIndex, ColumnIndex, RowIndex, ColumnIndex));
                                slws.SheetViews[i].Selections.Add(sel);
                            }
                        }
                        else
                        {
                            // else there are selections. We'll need to look for the selection that
                            // has TopLeft as the pane. Not sure if the Pane class is tightly connected
                            // to the Selection classes, so we're going to check separately.
                            // It appears that the Pane class exists only if the worksheet is split or
                            // frozen, but I might be wrong... And when the Pane class exists, then
                            // there seems to be 3 or 4 Selection classes. There seems to be 4 Selection
                            // classes only when the worksheet is split and the active cell is in the
                            // top left corner.
                            if (slws.SheetViews[i].HasPane)
                            {
                                vActivePane = slws.SheetViews[i].Pane.ActivePane;
                                for (j = slws.SheetViews[i].Selections.Count - 1; j >= 0; --j)
                                {
                                    if (slws.SheetViews[i].Selections[j].Pane == vActivePane)
                                    {
                                        slws.SheetViews[i].Selections[j].ActiveCell = SLTool.ToCellReference(RowIndex, ColumnIndex);
                                        slws.SheetViews[i].Selections[j].SequenceOfReferences.Clear();
                                        slws.SheetViews[i].Selections[j].SequenceOfReferences.Add(new SLCellPointRange(RowIndex, ColumnIndex, RowIndex, ColumnIndex));
                                    }
                                }
                            }
                            else
                            {
                                for (j = slws.SheetViews[i].Selections.Count - 1; j >= 0; --j)
                                {
                                    if (slws.SheetViews[i].Selections[j].Pane == PaneValues.TopLeft)
                                    {
                                        if (RowIndex == 1 && ColumnIndex == 1)
                                        {
                                            slws.SheetViews[i].Selections.RemoveAt(j);
                                        }
                                        else
                                        {
                                            slws.SheetViews[i].Selections[j].ActiveCell = SLTool.ToCellReference(RowIndex, ColumnIndex);
                                            slws.SheetViews[i].Selections[j].SequenceOfReferences.Clear();
                                            slws.SheetViews[i].Selections[j].SequenceOfReferences.Add(new SLCellPointRange(RowIndex, ColumnIndex, RowIndex, ColumnIndex));
                                        }
                                    }
                                }
                            }
                        }

                        break;
                    }
                }

                if (!bFound)
                {
                    sv = new SLSheetView();
                    sv.WorkbookViewId = 0;
                    sel = new SLSelection();
                    sel.ActiveCell = SLTool.ToCellReference(RowIndex, ColumnIndex);
                    sel.SequenceOfReferences.Add(new SLCellPointRange(RowIndex, ColumnIndex, RowIndex, ColumnIndex));
                    sv.Selections.Add(sel);

                    slws.SheetViews.Add(sv);
                }
            }

            return true;
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            return MergeWorksheetCellsFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, null, null);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            return MergeWorksheetCellsFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, null, b);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.Color = BorderColor;
            b.BottomBorder.Color = BorderColor;
            b.LeftBorder.Color = BorderColor;
            b.RightBorder.Color = BorderColor;

            return MergeWorksheetCellsFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, null, b);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor);
            b.BottomBorder.SetBorderThemeColor(BorderColor);
            b.LeftBorder.SetBorderThemeColor(BorderColor);
            b.RightBorder.SetBorderThemeColor(BorderColor);

            return MergeWorksheetCellsFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, null, b);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(string StartCellReference, string EndCellReference, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor, Tint);
            b.BottomBorder.SetBorderThemeColor(BorderColor, Tint);
            b.LeftBorder.SetBorderThemeColor(BorderColor, Tint);
            b.RightBorder.SetBorderThemeColor(BorderColor, Tint);

            return MergeWorksheetCellsFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, null, b);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        /// <param name="Border">The SLBorder object with border style properties.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(string StartCellReference, string EndCellReference, SLBorder Border)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            return MergeWorksheetCellsFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, null, Border);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Cell style and border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        /// <param name="Style">The SLStyle object with style properties. Any border style properties set in this SLStyle object will be used.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(string StartCellReference, string EndCellReference, SLStyle Style)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            return MergeWorksheetCellsFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Style, null);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return MergeWorksheetCellsFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, null, null);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            return MergeWorksheetCellsFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, null, b);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.Color = BorderColor;
            b.BottomBorder.Color = BorderColor;
            b.LeftBorder.Color = BorderColor;
            b.RightBorder.Color = BorderColor;

            return MergeWorksheetCellsFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, null, b);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor);
            b.BottomBorder.SetBorderThemeColor(BorderColor);
            b.LeftBorder.SetBorderThemeColor(BorderColor);
            b.RightBorder.SetBorderThemeColor(BorderColor);

            return MergeWorksheetCellsFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, null, b);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border theme color.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            SLBorder b = this.CreateBorder();
            b.TopBorder.BorderStyle = BorderStyle;
            b.BottomBorder.BorderStyle = BorderStyle;
            b.LeftBorder.BorderStyle = BorderStyle;
            b.RightBorder.BorderStyle = BorderStyle;

            b.TopBorder.SetBorderThemeColor(BorderColor, Tint);
            b.BottomBorder.SetBorderThemeColor(BorderColor, Tint);
            b.LeftBorder.SetBorderThemeColor(BorderColor, Tint);
            b.RightBorder.SetBorderThemeColor(BorderColor, Tint);

            return MergeWorksheetCellsFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, null, b);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        /// <param name="Border">The SLBorder object with border style properties.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLBorder Border)
        {
            return MergeWorksheetCellsFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, null, Border);
        }

        /// <summary>
        /// Merge cells given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell. No merging is done if it's just one cell. Cell style and border style properties are only applied on a successful merge.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        /// <param name="Style">The SLStyle object with style properties. Any border style properties set in this SLStyle object will be used.</param>
        /// <returns>True if merging is successful. False otherwise.</returns>
        public bool MergeWorksheetCells(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLStyle Style)
        {
            return MergeWorksheetCellsFinal(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, Style, null);
        }

        private bool MergeWorksheetCellsFinal(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, SLStyle Style, SLBorder Border)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            // no point merging one cell
            if (iStartRowIndex == iEndRowIndex && iStartColumnIndex == iEndColumnIndex)
            {
                return false;
            }

            int i;
            bool result = false;
            SLMergeCell mc = new SLMergeCell();
            if (SLTool.CheckRowColumnIndexLimit(iStartRowIndex, iStartColumnIndex) && SLTool.CheckRowColumnIndexLimit(iEndRowIndex, iEndColumnIndex))
            {
                result = true;
                for (i = 0; i < slws.MergeCells.Count; ++i)
                {
                    mc = slws.MergeCells[i];

                    // This comes from the separating axis theorem.
                    // We're checking that the given merged cell does not overlap with
                    // any existing merged cells. The conditions are made easier because
                    // the merged cells are rectangular, the row/column indices are whole numbers,
                    // and they map strictly to a 2D grid.
                    // We've also rearranged values such that the given end row index is equal
                    // to or greater than the given start row index (similarly for the column index).
                    // This means we only need to check for one given value against an existing value.

                    // The given merged cell doesn't overlap if:
                    // 1) it is completely above the existing merged cell OR
                    // 2) it is completely below the existing merged cell OR
                    // 3) it is completely to the left of the existing merged cell OR
                    // 4) it is completely to the right of the existing merged cell

                    if (!(iEndRowIndex < mc.StartRowIndex || iStartRowIndex > mc.EndRowIndex || iEndColumnIndex < mc.StartColumnIndex || iStartColumnIndex > mc.EndColumnIndex))
                    {
                        result = false;
                        break;
                    }
                }

                if (result)
                {
                    SLTable t;
                    for (i = 0; i < slws.Tables.Count; ++i)
                    {
                        t = slws.Tables[i];
                        if (!(iEndRowIndex < t.StartRowIndex || iStartRowIndex > t.EndRowIndex || iEndColumnIndex < t.StartColumnIndex || iStartColumnIndex > t.EndColumnIndex))
                        {
                            result = false;
                            break;
                        }
                    }
                }
            }

            // if all went well!
            if (result)
            {
                mc = new SLMergeCell();
                mc.FromIndices(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
                slws.MergeCells.Add(mc);

                if (Style != null)
                {
                    SLStyle cellstyle = Style.Clone();
                    // some optimisations. If the cell is the top-left and doesn't touch the right side,
                    // remove the right border. Similarly, if it doesn't touch the bottom side, remove
                    // the bottom border. Probably not much of an optimisation, because it's to reduce
                    // the number of unique styles held. Ah well, one more style probably won't kill Excel,
                    // but I try to be helpful when it's not too much trouble...
                    // This actually try to simulate what happens in DrawBorderFinal(). Basically, the
                    // top-left cell should be assigned the exact same border style when in DrawBorderFinal().
                    // This means by the powers that be, I mean, by the way we store the style hashes, it
                    // will be unique, so no extra border hash is created thus no extra style hash is created
                    // thus saving space! Thus optimisation! Thus world peace! (wait what?)
                    if (iStartColumnIndex != iEndColumnIndex) cellstyle.borderReal.RemoveRightBorder();
                    if (iStartRowIndex != iEndRowIndex) cellstyle.borderReal.RemoveBottomBorder();

                    // this is to facilitate things like centre-align and bold for the entire merge cell
                    this.SetCellStyle(iStartRowIndex, iStartColumnIndex, cellstyle);

                    // this is to facilitate things like border properties (duh)
                    // We assume the border property of the passed in SLStyle object is the one we use.
                    // We clone the passed in SLStyle object and then possibly remove the right and bottom
                    // border because the top-left cell doesn't necessarily have the right and bottom borders.
                    this.DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Style.Border, false);
                }
                else if (Border != null)
                {
                    this.DrawBorderFinal(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, Border, false);
                }
            }

            return result;
        }

        /// <summary>
        /// Unmerge cells given a corner cell of an existing merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        /// <returns>True if unmerging is successful. False otherwise.</returns>
        public bool UnmergeWorksheetCells(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            return UnmergeWorksheetCells(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Unmerge cells given a corner cell of an existing merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        /// <returns>True if unmerging is successful. False otherwise.</returns>
        public bool UnmergeWorksheetCells(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            bool result = false;
            SLMergeCell mc = new SLMergeCell();
            for (int i = 0; i < slws.MergeCells.Count; ++i)
            {
                mc = slws.MergeCells[i];
                if (mc.StartRowIndex == iStartRowIndex && mc.StartColumnIndex == iStartColumnIndex && mc.EndRowIndex == iEndRowIndex && mc.EndColumnIndex == iEndColumnIndex)
                {
                    slws.MergeCells.RemoveAt(i);
                    result = true;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Get a list of the existing merged cells.
        /// </summary>
        /// <returns>A list of the merged cells.</returns>
        public List<SLMergeCell> GetWorksheetMergeCells()
        {
            List<SLMergeCell> list = new List<SLMergeCell>();
            foreach (SLMergeCell mc in slws.MergeCells)
            {
                list.Add(mc.Clone());
            }

            return list;
        }

        /// <summary>
        /// Filter data.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the corner cell, such as "A1".</param>
        /// <param name="EndCellReference">The cell reference of the opposite corner cell, such as "A1".</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool Filter(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            return this.Filter(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Filter data.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool Filter(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            int i;
            bool result = false;
            if (SLTool.CheckRowColumnIndexLimit(iStartRowIndex, iStartColumnIndex) && SLTool.CheckRowColumnIndexLimit(iEndRowIndex, iEndColumnIndex))
            {
                result = true;

                // This comes from the separating axis theorem. See merging cells method for more details.

                // Technically, Excel allows you to filter a grid of cells with merged cells. But the
                // behaviour is a little dependent on the actual data. For example, you either select
                // the whole merged cell or you don't in the filter range. However, I'm not going to
                // enforce this.

                // Also technically speaking, you *can* filter a grid of cells that overlaps a table.
                // But the conditions are specific. The filter range must be completely within the table.
                // But the effect is that you remove the filter from the table. This is Excel! This is
                // because there's a visual interface.
                // So I'm going to assume the given filter range *cannot* overlap a table.
                SLTable t;
                for (i = 0; i < slws.Tables.Count; ++i)
                {
                    t = slws.Tables[i];
                    if (!(iEndRowIndex < t.StartRowIndex || iStartRowIndex > t.EndRowIndex || iEndColumnIndex < t.StartColumnIndex || iStartColumnIndex > t.EndColumnIndex))
                    {
                        result = false;
                        break;
                    }
                }

                if (result)
                {
                    slws.HasAutoFilter = true;
                    slws.AutoFilter = new SLAutoFilter();
                    slws.AutoFilter.StartRowIndex = iStartRowIndex;
                    slws.AutoFilter.StartColumnIndex = iStartColumnIndex;
                    slws.AutoFilter.EndRowIndex = iEndRowIndex;
                    slws.AutoFilter.EndColumnIndex = iEndColumnIndex;

                    int iLocalSheetID = -1;
                    for (i = 0; i < this.slwb.Sheets.Count; ++i)
                    {
                        if (this.slwb.Sheets[i].Name.Equals(this.gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            iLocalSheetID = i;
                            break;
                        }
                    }

                    if (iLocalSheetID >= 0)
                    {
                        bool bFound = false;
                        foreach (SLDefinedName dn in this.slwb.DefinedNames)
                        {
                            if (dn.LocalSheetId != null && dn.LocalSheetId.Value == (uint)iLocalSheetID && dn.Name.Equals("_xlnm._FilterDatabase", StringComparison.OrdinalIgnoreCase))
                            {
                                dn.Text = SLTool.ToCellRange(this.gsSelectedWorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, true);
                                dn.Hidden = true;
                                bFound = true;
                                break;
                            }
                        }

                        if (!bFound)
                        {
                            this.slwb.DefinedNames.Add(new SLDefinedName("_xlnm._FilterDatabase")
                            {
                                LocalSheetId = (uint)iLocalSheetID,
                                Hidden = true,
                                Text = SLTool.ToCellRange(this.gsSelectedWorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, true)
                            });
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Removing any data filter.
        /// </summary>
        public void RemoveFilter()
        {
            slws.HasAutoFilter = false;
            slws.AutoFilter = new SLAutoFilter();
        }

        /// <summary>
        /// Indicates if the currently selected worksheet has an existing filter.
        /// </summary>
        /// <returns>True if there's an existing filter. False otherwise.</returns>
        public bool HasFilter()
        {
            return slws.HasAutoFilter;
        }

        /// <summary>
        /// Get the filter range, if it exists. Call HasFilter() before calling this to make sure.
        /// </summary>
        /// <param name="StartRowIndex">The start row index of the filter range if it exists, and is -1 if it doesn't.</param>
        /// <param name="StartColumnIndex">The start column index of the filter range if it exists, and is -1 if it doesn't.</param>
        /// <param name="EndRowIndex">The end row index of the filter range if it exists, and is -1 if it doesn't.</param>
        /// <param name="EndColumnIndex">The end column index of the filter range if it exists, and is -1 if it doesn't.</param>
        public void GetFilterRange(ref int StartRowIndex, ref int StartColumnIndex, ref int EndRowIndex, ref int EndColumnIndex)
        {
            StartRowIndex = -1;
            StartColumnIndex = -1;
            EndRowIndex = -1;
            EndColumnIndex = -1;

            if (slws.HasAutoFilter)
            {
                StartRowIndex = slws.AutoFilter.StartRowIndex;
                StartColumnIndex = slws.AutoFilter.StartColumnIndex;
                EndRowIndex = slws.AutoFilter.EndRowIndex;
                EndColumnIndex = slws.AutoFilter.EndColumnIndex;
            }
        }

        /// <summary>
        /// Copy one cell to another cell.
        /// </summary>
        /// <param name="CellReference">The cell reference of the cell to be copied from, such as "A1".</param>
        /// <param name="AnchorCellReference">The cell reference of the cell to be copied to, such as "A1".</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(string CellReference, string AnchorCellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return CopyCell(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex, iAnchorRowIndex, iAnchorColumnIndex, false, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy one cell to another cell.
        /// </summary>
        /// <param name="CellReference">The cell reference of the cell to be copied from, such as "A1".</param>
        /// <param name="AnchorCellReference">The cell reference of the cell to be copied to, such as "A1".</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(string CellReference, string AnchorCellReference, bool ToCut)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return CopyCell(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex, iAnchorRowIndex, iAnchorColumnIndex, ToCut, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy one cell to another cell.
        /// </summary>
        /// <param name="CellReference">The cell reference of the cell to be copied from, such as "A1".</param>
        /// <param name="AnchorCellReference">The cell reference of the cell to be copied to, such as "A1".</param>
        /// <param name="PasteOption">Paste option.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(string CellReference, string AnchorCellReference, SLPasteTypeValues PasteOption)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return CopyCell(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex, iAnchorRowIndex, iAnchorColumnIndex, false, PasteOption);
        }

        /// <summary>
        /// Copy a range of cells to another range, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="AnchorCellReference">The cell reference of the anchor cell, such as "A1".</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(string StartCellReference, string EndCellReference, string AnchorCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return CopyCell(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, iAnchorRowIndex, iAnchorColumnIndex, false, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy a range of cells to another range, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="AnchorCellReference">The cell reference of the anchor cell, such as "A1".</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(string StartCellReference, string EndCellReference, string AnchorCellReference, bool ToCut)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return CopyCell(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, iAnchorRowIndex, iAnchorColumnIndex, ToCut, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy a range of cells to another range, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="AnchorCellReference">The cell reference of the anchor cell, such as "A1".</param>
        /// <param name="PasteOption">Paste options.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(string StartCellReference, string EndCellReference, string AnchorCellReference, SLPasteTypeValues PasteOption)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return CopyCell(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, iAnchorRowIndex, iAnchorColumnIndex, false, PasteOption);
        }

        /// <summary>
        /// Copy one cell to another cell.
        /// </summary>
        /// <param name="RowIndex">The row index of the cell to be copied from.</param>
        /// <param name="ColumnIndex">The column index of the cell to be copied from.</param>
        /// <param name="AnchorRowIndex">The row index of the cell to be copied to.</param>
        /// <param name="AnchorColumnIndex">The column index of the cell to be copied to.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(int RowIndex, int ColumnIndex, int AnchorRowIndex, int AnchorColumnIndex)
        {
            return CopyCell(RowIndex, ColumnIndex, RowIndex, ColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy one cell to another cell.
        /// </summary>
        /// <param name="RowIndex">The row index of the cell to be copied from.</param>
        /// <param name="ColumnIndex">The column index of the cell to be copied from.</param>
        /// <param name="AnchorRowIndex">The row index of the cell to be copied to.</param>
        /// <param name="AnchorColumnIndex">The column index of the cell to be copied to.</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(int RowIndex, int ColumnIndex, int AnchorRowIndex, int AnchorColumnIndex, bool ToCut)
        {
            return CopyCell(RowIndex, ColumnIndex, RowIndex, ColumnIndex, AnchorRowIndex, AnchorColumnIndex, ToCut, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy one cell to another cell.
        /// </summary>
        /// <param name="RowIndex">The row index of the cell to be copied from.</param>
        /// <param name="ColumnIndex">The column index of the cell to be copied from.</param>
        /// <param name="AnchorRowIndex">The row index of the cell to be copied to.</param>
        /// <param name="AnchorColumnIndex">The column index of the cell to be copied to.</param>
        /// <param name="PasteOption">Paste option.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(int RowIndex, int ColumnIndex, int AnchorRowIndex, int AnchorColumnIndex, SLPasteTypeValues PasteOption)
        {
            return CopyCell(RowIndex, ColumnIndex, RowIndex, ColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, PasteOption);
        }

        /// <summary>
        /// Copy a range of cells to another range, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor cell.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, int AnchorRowIndex, int AnchorColumnIndex)
        {
            return CopyCell(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy a range of cells to another range, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor cell.</param>
        /// <param name="ToCut">True for cut-and-paste. False for copy-and-paste.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, int AnchorRowIndex, int AnchorColumnIndex, bool ToCut)
        {
            return CopyCell(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, AnchorRowIndex, AnchorColumnIndex, ToCut, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy a range of cells to another range, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor cell.</param>
        /// <param name="PasteOption">Paste option.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCell(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, int AnchorRowIndex, int AnchorColumnIndex, SLPasteTypeValues PasteOption)
        {
            return CopyCell(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, PasteOption);
        }

        private bool CopyCell(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, int AnchorRowIndex, int AnchorColumnIndex, bool ToCut, SLPasteTypeValues PasteOption)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            bool result = false;
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit
                && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit
                && iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit
                && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit
                && AnchorRowIndex >= 1 && AnchorRowIndex <= SLConstants.RowLimit
                && AnchorColumnIndex >= 1 && AnchorColumnIndex <= SLConstants.ColumnLimit
                && (iStartRowIndex != AnchorRowIndex || iStartColumnIndex != AnchorColumnIndex))
            {
                this.FlattenAllSharedCellFormula();

                result = true;

                int i, j, iSwap, iStyleIndex, iStyleIndexNew;
                SLCell origcell, newcell;
                int iRowIndex, iColumnIndex;
                int iNewRowIndex, iNewColumnIndex;
                int rowdiff = AnchorRowIndex - iStartRowIndex;
                int coldiff = AnchorColumnIndex - iStartColumnIndex;
                SLCellWarehouse cells = new SLCellWarehouse();

                Dictionary<int, uint> colstyleindex = new Dictionary<int, uint>();
                Dictionary<int, uint> rowstyleindex = new Dictionary<int, uint>();

                List<int> rowindexkeys = slws.RowProperties.Keys.ToList<int>();
                SLRowProperties rp;
                foreach (int rowindex in rowindexkeys)
                {
                    rp = slws.RowProperties[rowindex];
                    rowstyleindex[rowindex] = rp.StyleIndex;
                }

                List<int> colindexkeys = slws.ColumnProperties.Keys.ToList<int>();
                SLColumnProperties cp;
                foreach (int colindex in colindexkeys)
                {
                    cp = slws.ColumnProperties[colindex];
                    colstyleindex[colindex] = cp.StyleIndex;
                }

                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        iRowIndex = i;
                        iColumnIndex = j;
                        iNewRowIndex = i + rowdiff;
                        iNewColumnIndex = j + coldiff;
                        if (ToCut)
                        {
                            if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                            {
                                cells.SetValue(iNewRowIndex, iNewColumnIndex, slws.CellWarehouse.Cells[iRowIndex][iColumnIndex]);
                                slws.CellWarehouse.Remove(iRowIndex, iColumnIndex);
                            }
                        }
                        else
                        {
                            switch (PasteOption)
                            {
                                case SLPasteTypeValues.Formatting:
                                    if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                                    {
                                        origcell = slws.CellWarehouse.Cells[iRowIndex][iColumnIndex];
                                        if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                        {
                                            newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                            newcell.StyleIndex = origcell.StyleIndex;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                        else
                                        {
                                            if (origcell.StyleIndex != 0)
                                            {
                                                // if not the default style, then must create a new
                                                // destination cell.
                                                newcell = new SLCell();
                                                newcell.StyleIndex = origcell.StyleIndex;
                                                newcell.CellText = string.Empty;
                                                cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                            }
                                            else
                                            {
                                                // else source cell has default style.
                                                // Now check if destination cell lies on a row/column
                                                // that has non-default style. Remember, we don't have 
                                                // a destination cell here.
                                                iStyleIndexNew = 0;
                                                if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                                if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                                if (iStyleIndexNew != 0)
                                                {
                                                    newcell = new SLCell();
                                                    newcell.StyleIndex = 0;
                                                    newcell.CellText = string.Empty;
                                                    cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // else no source cell
                                        if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                        {
                                            iStyleIndex = 0;
                                            if (rowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)rowstyleindex[iRowIndex];
                                            if (iStyleIndex == 0 && colstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)colstyleindex[iColumnIndex];

                                            newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                            newcell.StyleIndex = (uint)iStyleIndex;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                        else
                                        {
                                            // else no source and no destination, so we check for row/column
                                            // with non-default styles.
                                            iStyleIndex = 0;
                                            if (rowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)rowstyleindex[iRowIndex];
                                            if (iStyleIndex == 0 && colstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)colstyleindex[iColumnIndex];

                                            iStyleIndexNew = 0;
                                            if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                            if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                            if (iStyleIndex != 0 || iStyleIndexNew != 0)
                                            {
                                                newcell = new SLCell();
                                                newcell.StyleIndex = (uint)iStyleIndex;
                                                newcell.CellText = string.Empty;
                                                cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                            }
                                        }
                                    }
                                    break;
                                case SLPasteTypeValues.Formulas:
                                    if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                                    {
                                        origcell = slws.CellWarehouse.Cells[iRowIndex][iColumnIndex];
                                        if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                        {
                                            newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                            if (origcell.CellFormula != null) newcell.CellFormula = origcell.CellFormula.Clone();
                                            else newcell.CellFormula = null;
                                            newcell.CellText = origcell.CellText;
                                            newcell.fNumericValue = origcell.fNumericValue;
                                            newcell.DataType = origcell.DataType;
                                            ProcessCellFormulaDelta(ref newcell, false, iStartRowIndex, iStartColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, false, false, 0, 0);
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                        else
                                        {
                                            newcell = new SLCell();
                                            if (origcell.CellFormula != null) newcell.CellFormula = origcell.CellFormula.Clone();
                                            else newcell.CellFormula = null;
                                            newcell.CellText = origcell.CellText;
                                            newcell.fNumericValue = origcell.fNumericValue;
                                            newcell.DataType = origcell.DataType;

                                            iStyleIndexNew = 0;
                                            if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                            if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                            if (iStyleIndexNew != 0) newcell.StyleIndex = (uint)iStyleIndexNew;
                                            ProcessCellFormulaDelta(ref newcell, false, iStartRowIndex, iStartColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, false, false, 0, 0);
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                    }
                                    else
                                    {
                                        if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                        {
                                            newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                            newcell.CellText = string.Empty;
                                            newcell.DataType = CellValues.Number;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                        // no else because don't have to do anything
                                    }
                                    break;
                                case SLPasteTypeValues.Paste:
                                    if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                                    {
                                        origcell = slws.CellWarehouse.Cells[iRowIndex][iColumnIndex].Clone();
                                        newcell = origcell.Clone();
                                        ProcessCellFormulaDelta(ref newcell, false, iStartRowIndex, iStartColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, false, false, 0, 0);
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                    }
                                    else
                                    {
                                        // else the source cell is empty
                                        if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                        {
                                            iStyleIndex = 0;
                                            if (rowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)rowstyleindex[iRowIndex];
                                            if (iStyleIndex == 0 && colstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)colstyleindex[iColumnIndex];

                                            if (iStyleIndex != 0)
                                            {
                                                newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                                newcell.StyleIndex = (uint)iStyleIndex;
                                                newcell.CellText = string.Empty;
                                                cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                            }
                                            else
                                            {
                                                // if the source cell is empty, then direct pasting
                                                // means overwrite the existing cell, which is faster
                                                // by just removing it.
                                                slws.CellWarehouse.Remove(iNewRowIndex, iNewColumnIndex);
                                            }
                                        }
                                        else
                                        {
                                            // else no source and no destination, so we check for row/column
                                            // with non-default styles.
                                            iStyleIndex = 0;
                                            if (rowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)rowstyleindex[iRowIndex];
                                            if (iStyleIndex == 0 && colstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)colstyleindex[iColumnIndex];

                                            iStyleIndexNew = 0;
                                            if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                            if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                            if (iStyleIndex != 0 || iStyleIndexNew != 0)
                                            {
                                                newcell = new SLCell();
                                                newcell.StyleIndex = (uint)iStyleIndex;
                                                newcell.CellText = string.Empty;
                                                cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                            }
                                        }
                                    }
                                    break;
                                case SLPasteTypeValues.Transpose:
                                    iNewRowIndex = i - iStartRowIndex;
                                    iNewColumnIndex = j - iStartColumnIndex;
                                    iSwap = iNewRowIndex;
                                    iNewRowIndex = iNewColumnIndex;
                                    iNewColumnIndex = iSwap;
                                    iNewRowIndex = iNewRowIndex + iStartRowIndex + rowdiff;
                                    iNewColumnIndex = iNewColumnIndex + iStartColumnIndex + coldiff;
                                    // in case say the millionth row is transposed, because we can't have a millionth column.
                                    if (iNewRowIndex <= SLConstants.RowLimit && iNewColumnIndex <= SLConstants.ColumnLimit)
                                    {
                                        // this part is identical to normal paste except for the formula processing.
                                        // We swap the row and column diff's because it's a transpose.

                                        if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                                        {
                                            origcell = slws.CellWarehouse.Cells[iRowIndex][iColumnIndex].Clone();
                                            ProcessCellFormulaDelta(ref origcell, true, i, j, iNewRowIndex, iNewColumnIndex, false, false, false, 0, 0);
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, origcell);
                                        }
                                        else
                                        {
                                            // else the source cell is empty
                                            if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                            {
                                                iStyleIndex = 0;
                                                if (rowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)rowstyleindex[iRowIndex];
                                                if (iStyleIndex == 0 && colstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)colstyleindex[iColumnIndex];

                                                if (iStyleIndex != 0)
                                                {
                                                    newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                                    newcell.StyleIndex = (uint)iStyleIndex;
                                                    newcell.CellText = string.Empty;
                                                    cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                                }
                                                else
                                                {
                                                    // if the source cell is empty, then direct pasting
                                                    // means overwrite the existing cell, which is faster
                                                    // by just removing it.
                                                    slws.CellWarehouse.Remove(iNewRowIndex, iNewColumnIndex);
                                                }
                                            }
                                            else
                                            {
                                                // else no source and no destination, so we check for row/column
                                                // with non-default styles.
                                                iStyleIndex = 0;
                                                if (rowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)rowstyleindex[iRowIndex];
                                                if (iStyleIndex == 0 && colstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)colstyleindex[iColumnIndex];

                                                iStyleIndexNew = 0;
                                                if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                                if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                                if (iStyleIndex != 0 || iStyleIndexNew != 0)
                                                {
                                                    newcell = new SLCell();
                                                    newcell.StyleIndex = (uint)iStyleIndex;
                                                    newcell.CellText = string.Empty;
                                                    cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                                }
                                            }
                                        }
                                    }
                                    break;
                                case SLPasteTypeValues.Values:
                                    // this part is identical to the formula part, except
                                    // for assigning the cell formula part.

                                    if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                                    {
                                        origcell = slws.CellWarehouse.Cells[iRowIndex][iColumnIndex];
                                        if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                        {
                                            newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                            newcell.CellFormula = null;
                                            newcell.CellText = origcell.CellText;
                                            newcell.fNumericValue = origcell.fNumericValue;
                                            newcell.DataType = origcell.DataType;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                        else
                                        {
                                            newcell = new SLCell();
                                            newcell.CellFormula = null;
                                            newcell.CellText = origcell.CellText;
                                            newcell.fNumericValue = origcell.fNumericValue;
                                            newcell.DataType = origcell.DataType;
                                            
                                            iStyleIndexNew = 0;
                                            if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                            if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                            if (iStyleIndexNew != 0) newcell.StyleIndex = (uint)iStyleIndexNew;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                    }
                                    else
                                    {
                                        if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                        {
                                            newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                            newcell.CellFormula = null;
                                            newcell.CellText = string.Empty;
                                            newcell.DataType = CellValues.Number;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                        // no else because don't have to do anything
                                    }
                                    break;
                            }
                        }
                    }
                }

                int AnchorEndRowIndex = AnchorRowIndex + iEndRowIndex - iStartRowIndex;
                int AnchorEndColumnIndex = AnchorColumnIndex + iEndColumnIndex - iStartColumnIndex;

                for (i = AnchorRowIndex; i <= AnchorEndRowIndex; ++i)
                {
                    for (j = AnchorColumnIndex; j <= AnchorEndColumnIndex; ++j)
                    {
                        iRowIndex = i;
                        iColumnIndex = j;
                        // any cell within destination "paste" operation is taken out
                        slws.CellWarehouse.Remove(iRowIndex, iColumnIndex);
                    }
                }

                int iNumberOfRows = iEndRowIndex - iStartRowIndex + 1;
                if (AnchorRowIndex <= iStartRowIndex) iNumberOfRows = -iNumberOfRows;
                int iNumberOfColumns = iEndColumnIndex - iStartColumnIndex + 1;
                if (AnchorColumnIndex <= iStartColumnIndex) iNumberOfColumns = -iNumberOfColumns;
                List<int> listRowKeys = cells.Cells.Keys.ToList<int>();
                List<int> listColumnKeys;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = cells.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        origcell = cells.Cells[rowkey][colkey];
                        //if (PasteOption != SLPasteTypeValues.Transpose)
                        //{
                        //    //this.ProcessCellFormulaDelta(ref origcell, AnchorRowIndex, iNumberOfRows, AnchorColumnIndex, iNumberOfColumns);
                        //    this.ProcessCellFormulaDelta(ref origcell, iNumberOfRows, iNumberOfColumns);
                        //}
                        //else
                        //{
                        //    //this.ProcessCellFormulaDelta(ref origcell, AnchorRowIndex, iNumberOfColumns, AnchorColumnIndex, iNumberOfRows);
                        //    this.ProcessCellFormulaDelta(ref origcell, iNumberOfColumns, iNumberOfRows);
                        //}
                        slws.CellWarehouse.SetValue(rowkey, colkey, origcell);
                    }
                }

                // TODO: tables!

                // cutting and pasting into a region with merged cells unmerges the existing merged cells
                // copying and pasting into a region with merged cells leaves existing merged cells alone.
                // Why does Excel do that? Don't know.
                // Will just standardise to leaving existing merged cells alone.
                List<SLMergeCell> mca = this.GetWorksheetMergeCells();
                foreach (SLMergeCell mc in mca)
                {
                    if (mc.StartRowIndex >= iStartRowIndex && mc.EndRowIndex <= iEndRowIndex
                        && mc.StartColumnIndex >= iStartColumnIndex && mc.EndColumnIndex <= iEndColumnIndex)
                    {
                        if (ToCut)
                        {
                            slws.MergeCells.Remove(mc);
                        }

                        if (PasteOption == SLPasteTypeValues.Transpose)
                        {
                            iRowIndex = mc.StartRowIndex - iStartRowIndex;
                            iColumnIndex = mc.StartColumnIndex - iStartColumnIndex;
                            iSwap = iRowIndex;
                            iRowIndex = iColumnIndex;
                            iColumnIndex = iSwap;
                            iRowIndex = iRowIndex + iStartRowIndex + rowdiff;
                            iColumnIndex = iColumnIndex + iStartColumnIndex + coldiff;

                            iNewRowIndex = mc.EndRowIndex - iStartRowIndex;
                            iNewColumnIndex = mc.EndColumnIndex - iStartColumnIndex;
                            iSwap = iNewRowIndex;
                            iNewRowIndex = iNewColumnIndex;
                            iNewColumnIndex = iSwap;
                            iNewRowIndex = iNewRowIndex + iStartRowIndex + rowdiff;
                            iNewColumnIndex = iNewColumnIndex + iStartColumnIndex + coldiff;

                            this.MergeWorksheetCells(iRowIndex, iColumnIndex, iNewRowIndex, iNewColumnIndex);
                        }
                        else
                        {
                            this.MergeWorksheetCells(mc.StartRowIndex + rowdiff, mc.StartColumnIndex + coldiff, mc.EndRowIndex + rowdiff, mc.EndColumnIndex + coldiff);
                        }
                    }
                }

                // TODO: conditional formatting and data validations?

                #region Hyperlinks
                if (slws.Hyperlinks.Count > 0)
                {
                    if (ToCut)
                    {
                        foreach (SLHyperlink hl in slws.Hyperlinks)
                        {
                            // if hyperlink is completely within copy range
                            if (iStartRowIndex <= hl.Reference.StartRowIndex
                                && hl.Reference.EndRowIndex <= iEndRowIndex
                                && iStartColumnIndex <= hl.Reference.StartColumnIndex
                                && hl.Reference.EndColumnIndex <= iEndColumnIndex)
                            {
                                hl.Reference = new SLCellPointRange(hl.Reference.StartRowIndex + rowdiff,
                                    hl.Reference.StartColumnIndex + coldiff,
                                    hl.Reference.EndRowIndex + rowdiff,
                                    hl.Reference.EndColumnIndex + coldiff);
                            }
                            // else don't change anything (Excel doesn't, so we don't).
                        }
                    }
                    else
                    {
                        // we only care if normal paste or transpose paste. Just like Excel.
                        if (PasteOption == SLPasteTypeValues.Paste || PasteOption == SLPasteTypeValues.Transpose)
                        {
                            List<SLHyperlink> copiedhyperlinks = new List<SLHyperlink>();
                            SLHyperlink hlCopied;

                            // hyperlink ID, URL
                            Dictionary<string, string> hlurl = new Dictionary<string, string>();

                            if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
                            {
                                WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                                foreach (HyperlinkRelationship hlrel in wsp.HyperlinkRelationships)
                                {
                                    if (hlrel.IsExternal)
                                    {
                                        hlurl[hlrel.Id] = hlrel.Uri.OriginalString;
                                    }
                                }
                            }

                            int iOverlapStartRowIndex = 1;
                            int iOverlapStartColumnIndex = 1;
                            int iOverlapEndRowIndex = 1;
                            int iOverlapEndColumnIndex = 1;
                            foreach (SLHyperlink hl in slws.Hyperlinks)
                            {
                                // this comes from the separating axis theorem.
                                // See merged cells for more details.
                                // In this case however, we're doing stuff when there's overlapping.
                                if (!(iEndRowIndex < hl.Reference.StartRowIndex
                                    || iStartRowIndex > hl.Reference.EndRowIndex
                                    || iEndColumnIndex < hl.Reference.StartColumnIndex
                                    || iStartColumnIndex > hl.Reference.EndColumnIndex))
                                {
                                    // get the overlapping region
                                    iOverlapStartRowIndex = Math.Max(iStartRowIndex, hl.Reference.StartRowIndex);
                                    iOverlapStartColumnIndex = Math.Max(iStartColumnIndex, hl.Reference.StartColumnIndex);
                                    iOverlapEndRowIndex = Math.Min(iEndRowIndex, hl.Reference.EndRowIndex);
                                    iOverlapEndColumnIndex = Math.Min(iEndColumnIndex, hl.Reference.EndColumnIndex);

                                    // offset to the correctly pasted region
                                    if (PasteOption == SLPasteTypeValues.Paste)
                                    {
                                        iOverlapStartRowIndex += rowdiff;
                                        iOverlapStartColumnIndex += coldiff;
                                        iOverlapEndRowIndex += rowdiff;
                                        iOverlapEndColumnIndex += coldiff;
                                    }
                                    else
                                    {
                                        // can only be transpose. See if check above.

                                        if (iOverlapEndRowIndex > SLConstants.ColumnLimit)
                                        {
                                            // probably won't happen. This means that after transpose,
                                            // the end row index will flip to exceed the column limit.
                                            // I don't feel like testing how Excel handles this, so
                                            // I'm going to just take it as normal paste.
                                            iOverlapStartRowIndex += rowdiff;
                                            iOverlapStartColumnIndex += coldiff;
                                            iOverlapEndRowIndex += rowdiff;
                                            iOverlapEndColumnIndex += coldiff;
                                        }
                                        else
                                        {
                                            iOverlapStartRowIndex -= iStartRowIndex;
                                            iOverlapStartColumnIndex -= iStartColumnIndex;
                                            iOverlapEndRowIndex -= iStartRowIndex;
                                            iOverlapEndColumnIndex -= iStartColumnIndex;

                                            iSwap = iOverlapStartRowIndex;
                                            iOverlapStartRowIndex = iOverlapStartColumnIndex;
                                            iOverlapStartColumnIndex = iSwap;

                                            iSwap = iOverlapEndRowIndex;
                                            iOverlapEndRowIndex = iOverlapEndColumnIndex;
                                            iOverlapEndColumnIndex = iSwap;

                                            iOverlapStartRowIndex += (iStartRowIndex + rowdiff);
                                            iOverlapStartColumnIndex += (iStartColumnIndex + coldiff);
                                            iOverlapEndRowIndex += (iStartRowIndex + rowdiff);
                                            iOverlapEndColumnIndex += (iStartColumnIndex + coldiff);
                                        }
                                    }

                                    hlCopied = new SLHyperlink();
                                    hlCopied = hl.Clone();
                                    hlCopied.IsNew = true;
                                    if (hlCopied.IsExternal)
                                    {
                                        if (hlurl.ContainsKey(hlCopied.Id))
                                        {
                                            hlCopied.HyperlinkUri = hlurl[hlCopied.Id];
                                            if (hlCopied.HyperlinkUri.StartsWith("."))
                                            {
                                                // assume this is a relative file path such as ../ or ./
                                                hlCopied.HyperlinkUriKind = UriKind.Relative;
                                            }
                                            else
                                            {
                                                hlCopied.HyperlinkUriKind = UriKind.Absolute;
                                            }
                                            hlCopied.Id = string.Empty;
                                        }
                                    }
                                    hlCopied.Reference = new SLCellPointRange(iOverlapStartRowIndex, iOverlapStartColumnIndex, iOverlapEndRowIndex, iOverlapEndColumnIndex);
                                    copiedhyperlinks.Add(hlCopied);
                                }
                            }

                            if (copiedhyperlinks.Count > 0)
                            {
                                slws.Hyperlinks.AddRange(copiedhyperlinks);
                            }
                        }
                    }
                }
                #endregion

                #region Calculation cells
                if (slwb.CalculationCells.Count > 0)
                {
                    // I don't know enough to fiddle with calculation chains. So I'm going to ignore it.
                    slwb.CalculationCells.Clear();

                    //List<int> listToDelete = new List<int>();
                    //int iRowIndex = -1;
                    //int iColumnIndex = -1;
                    //for (i = 0; i < slwb.CalculationCells.Count; ++i)
                    //{
                    //    if (slwb.CalculationCells[i].SheetId == giSelectedWorksheetID)
                    //    {
                    //        iRowIndex = slwb.CalculationCells[i].RowIndex;
                    //        iColumnIndex = slwb.CalculationCells[i].ColumnIndex;
                    //        if (ToCut && iRowIndex >= iStartRowIndex && iRowIndex <= iEndRowIndex
                    //                && iColumnIndex >= iStartColumnIndex && iColumnIndex <= iEndColumnIndex)
                    //        {
                    //            // just remove because recalculation of cell references is too complicated...
                    //            if (!listToDelete.Contains(i)) listToDelete.Add(i);
                    //        }

                    //        if (iRowIndex >= AnchorRowIndex && iRowIndex <= AnchorEndRowIndex
                    //            && iColumnIndex >= AnchorColumnIndex && iColumnIndex <= AnchorEndColumnIndex)
                    //        {
                    //            // existing calculation cell lies within destination "paste" operation
                    //            if (!listToDelete.Contains(i)) listToDelete.Add(i);
                    //        }
                    //    }
                    //}

                    //for (i = listToDelete.Count - 1; i >= 0; --i)
                    //{
                    //    slwb.CalculationCells.RemoveAt(listToDelete[i]);
                    //}
                }
                #endregion

                // defined names is hard to calculate...
                // need to check the row and column indices based on the cell references within.
            }

            return result;
        }

        /// <summary>
        /// Copy one cell from another worksheet to the currently selected worksheet.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="CellReference">The cell reference of the cell to be copied from, such as "A1".</param>
        /// <param name="AnchorCellReference">The cell reference of the anchor cell, such as "A1".</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellFromWorksheet(string WorksheetName, string CellReference, string AnchorCellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return this.CopyCellFromWorksheet(WorksheetName, iRowIndex, iColumnIndex, iRowIndex, iColumnIndex, iAnchorRowIndex, iAnchorColumnIndex, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy one cell from another worksheet to the currently selected worksheet.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="CellReference">The cell reference of the cell to be copied from, such as "A1".</param>
        /// <param name="AnchorCellReference">The cell reference of the anchor cell, such as "A1".</param>
        /// <param name="PasteOption">Paste option.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellFromWorksheet(string WorksheetName, string CellReference, string AnchorCellReference, SLPasteTypeValues PasteOption)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return this.CopyCellFromWorksheet(WorksheetName, iRowIndex, iColumnIndex, iRowIndex, iColumnIndex, iAnchorRowIndex, iAnchorColumnIndex, PasteOption);
        }

        /// <summary>
        /// Copy a range of cells from another worksheet to the currently selected worksheet, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="AnchorCellReference">The cell reference of the anchor cell, such as "A1".</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellFromWorksheet(string WorksheetName, string StartCellReference, string EndCellReference, string AnchorCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return this.CopyCellFromWorksheet(WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, iAnchorRowIndex, iAnchorColumnIndex);
        }

        /// <summary>
        /// Copy a range of cells from another worksheet to the currently selected worksheet, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range, such as "A1". This is typically the bottom-right cell.</param>
        /// <param name="AnchorCellReference">The cell reference of the anchor cell, such as "A1".</param>
        /// <param name="PasteOption">Paste option.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellFromWorksheet(string WorksheetName, string StartCellReference, string EndCellReference, string AnchorCellReference, SLPasteTypeValues PasteOption)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            int iAnchorRowIndex = -1;
            int iAnchorColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(AnchorCellReference, out iAnchorRowIndex, out iAnchorColumnIndex))
            {
                return false;
            }

            return this.CopyCellFromWorksheet(WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, iAnchorRowIndex, iAnchorColumnIndex, PasteOption);
        }

        /// <summary>
        /// Copy one cell from another worksheet to the currently selected worksheet.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="RowIndex">The row index of the cell to be copied from.</param>
        /// <param name="ColumnIndex">The column index of the cell to be copied from.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor cell.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellFromWorksheet(string WorksheetName, int RowIndex, int ColumnIndex, int AnchorRowIndex, int AnchorColumnIndex)
        {
            return this.CopyCellFromWorksheet(WorksheetName, RowIndex, ColumnIndex, RowIndex, ColumnIndex, AnchorRowIndex, AnchorColumnIndex);
        }

        /// <summary>
        /// Copy one cell from another worksheet to the currently selected worksheet.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="RowIndex">The row index of the cell to be copied from.</param>
        /// <param name="ColumnIndex">The column index of the cell to be copied from.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor cell.</param>
        /// <param name="PasteOption">Paste option.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellFromWorksheet(string WorksheetName, int RowIndex, int ColumnIndex, int AnchorRowIndex, int AnchorColumnIndex, SLPasteTypeValues PasteOption)
        {
            return this.CopyCellFromWorksheet(WorksheetName, RowIndex, ColumnIndex, RowIndex, ColumnIndex, AnchorRowIndex, AnchorColumnIndex, PasteOption);
        }

        /// <summary>
        /// Copy a range of cells from another worksheet to the currently selected worksheet, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor cell.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellFromWorksheet(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, int AnchorRowIndex, int AnchorColumnIndex)
        {
            return this.CopyCellFromWorksheet(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, AnchorRowIndex, AnchorColumnIndex, SLPasteTypeValues.Paste);
        }

        /// <summary>
        /// Copy a range of cells from another worksheet to the currently selected worksheet, given the anchor cell of the destination range (top-left cell).
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="AnchorRowIndex">The row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">The column index of the anchor cell.</param>
        /// <param name="PasteOption">Paste option.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool CopyCellFromWorksheet(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, int AnchorRowIndex, int AnchorColumnIndex, SLPasteTypeValues PasteOption)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            if (WorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                return this.CopyCell(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, AnchorRowIndex, AnchorColumnIndex, false);
            }

            string sRelId = string.Empty;
            foreach (SLSheet sheet in slwb.Sheets)
            {
                if (sheet.Name.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    sRelId = sheet.Id;
                    break;
                }
            }

            // there has to be a valid existing worksheet
            if (sRelId.Length == 0) return false;

            bool result = false;
            if (iStartRowIndex >= 1 && iStartRowIndex <= SLConstants.RowLimit
                && iEndRowIndex >= 1 && iEndRowIndex <= SLConstants.RowLimit
                && iStartColumnIndex >= 1 && iStartColumnIndex <= SLConstants.ColumnLimit
                && iEndColumnIndex >= 1 && iEndColumnIndex <= SLConstants.ColumnLimit
                && AnchorRowIndex >= 1 && AnchorRowIndex <= SLConstants.RowLimit
                && AnchorColumnIndex >= 1 && AnchorColumnIndex <= SLConstants.ColumnLimit)
            {
                this.FlattenAllSharedCellFormula();

                result = true;

                WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(sRelId);

                int i, j, iSwap, iStyleIndex, iStyleIndexNew;
                SLCell origcell, newcell;
                int iRowIndex, iColumnIndex;
                int iNewRowIndex, iNewColumnIndex;
                int rowdiff = AnchorRowIndex - iStartRowIndex;
                int coldiff = AnchorColumnIndex - iStartColumnIndex;
                SLCellWarehouse cells = new SLCellWarehouse();
                SLCellWarehouse sourcecells = new SLCellWarehouse();

                Dictionary<int, uint> sourcecolstyleindex = new Dictionary<int, uint>();
                Dictionary<int, uint> sourcerowstyleindex = new Dictionary<int, uint>();

                string sCellRef = string.Empty;
                Dictionary<int, HashSet<int>> multiCellRef = new Dictionary<int, HashSet<int>>();

                // I use a hash set on the logic that it's easier to check a dictionary/hash
                // first, rather than load a Cell class into SLCell and then check with row/column indices.
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    multiCellRef.Add(i, new HashSet<int>());
                    for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        multiCellRef[i].Add(j);
                    }
                }

                // hyperlink ID, URL
                Dictionary<string, string> hlurl = new Dictionary<string, string>();
                List<SLHyperlink> sourcehyperlinks = new List<SLHyperlink>();

                foreach (HyperlinkRelationship hlrel in wsp.HyperlinkRelationships)
                {
                    if (hlrel.IsExternal)
                    {
                        hlurl[hlrel.Id] = hlrel.Uri.OriginalString;
                    }
                }

                using (OpenXmlReader oxr = OpenXmlReader.Create(wsp))
                {
                    Column col;
                    int iColumnMin, iColumnMax;
                    Row r;
                    Cell c;
                    SLHyperlink hl;
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(Column))
                        {
                            col = (Column)oxr.LoadCurrentElement();
                            iColumnMin = (int)col.Min.Value;
                            iColumnMax = (int)col.Max.Value;
                            for (i = iColumnMin; i <= iColumnMax; ++i)
                            {
                                sourcecolstyleindex[i] = (col.Style != null) ? col.Style.Value : 0;
                            }
                        }
                        else if (oxr.ElementType == typeof(Row))
                        {
                            r = (Row)oxr.LoadCurrentElement();
                            if (r.RowIndex != null)
                            {
                                if (r.StyleIndex != null) sourcerowstyleindex[(int)r.RowIndex.Value] = r.StyleIndex.Value;
                                else sourcerowstyleindex[(int)r.RowIndex.Value] = 0;
                            }

                            using (OpenXmlReader oxrRow = OpenXmlReader.Create(r))
                            {
                                while (oxrRow.Read())
                                {
                                    if (oxrRow.ElementType == typeof(Cell))
                                    {
                                        c = (Cell)oxrRow.LoadCurrentElement();
                                        if (c.CellReference != null)
                                        {
                                            sCellRef = c.CellReference.Value;
                                            if (SLTool.FormatCellReferenceToRowColumnIndex(sCellRef, out i, out j))
                                            {
                                                if (multiCellRef.ContainsKey(i) && multiCellRef[i].Contains(j))
                                                {
                                                    origcell = new SLCell();
                                                    origcell.FromCell(c);
                                                    sourcecells.SetValue(i, j, origcell);
                                                }
                                            }

                                        }
                                    }
                                }
                            }
                        }
                        else if (oxr.ElementType == typeof(Hyperlink))
                        {
                            hl = new SLHyperlink();
                            hl.FromHyperlink((Hyperlink)oxr.LoadCurrentElement());
                            sourcehyperlinks.Add(hl);
                        }
                    }
                }

                Dictionary<int, uint> colstyleindex = new Dictionary<int, uint>();
                Dictionary<int, uint> rowstyleindex = new Dictionary<int, uint>();

                List<int> rowindexkeys = slws.RowProperties.Keys.ToList<int>();
                SLRowProperties rp;
                foreach (int rowindex in rowindexkeys)
                {
                    rp = slws.RowProperties[rowindex];
                    rowstyleindex[rowindex] = rp.StyleIndex;
                }

                List<int> colindexkeys = slws.ColumnProperties.Keys.ToList<int>();
                SLColumnProperties cp;
                foreach (int colindex in colindexkeys)
                {
                    cp = slws.ColumnProperties[colindex];
                    colstyleindex[colindex] = cp.StyleIndex;
                }

                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        iRowIndex = i;
                        iColumnIndex = j;
                        iNewRowIndex = i + rowdiff;
                        iNewColumnIndex = j + coldiff;
                        switch (PasteOption)
                        {
                            case SLPasteTypeValues.Formatting:
                                if (sourcecells.Exists(iRowIndex, iColumnIndex))
                                {
                                    origcell = sourcecells.Cells[iRowIndex][iColumnIndex];
                                    if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                    {
                                        newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                        newcell.StyleIndex = origcell.StyleIndex;
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                    }
                                    else
                                    {
                                        if (origcell.StyleIndex != 0)
                                        {
                                            // if not the default style, then must create a new
                                            // destination cell.
                                            newcell = new SLCell();
                                            newcell.StyleIndex = origcell.StyleIndex;
                                            newcell.CellText = string.Empty;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                        else
                                        {
                                            // else source cell has default style.
                                            // Now check if destination cell lies on a row/column
                                            // that has non-default style. Remember, we don't have 
                                            // a destination cell here.
                                            iStyleIndexNew = 0;
                                            if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                            if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                            if (iStyleIndexNew != 0)
                                            {
                                                newcell = new SLCell();
                                                newcell.StyleIndex = 0;
                                                newcell.CellText = string.Empty;
                                                cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    // else no source cell
                                    if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                    {
                                        iStyleIndex = 0;
                                        if (sourcerowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)sourcerowstyleindex[iRowIndex];
                                        if (iStyleIndex == 0 && sourcecolstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)sourcecolstyleindex[iColumnIndex];

                                        newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                        newcell.StyleIndex = (uint)iStyleIndex;
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                    }
                                    else
                                    {
                                        // else no source and no destination, so we check for row/column
                                        // with non-default styles.
                                        iStyleIndex = 0;
                                        if (sourcerowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)sourcerowstyleindex[iRowIndex];
                                        if (iStyleIndex == 0 && sourcecolstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)sourcecolstyleindex[iColumnIndex];

                                        iStyleIndexNew = 0;
                                        if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                        if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                        if (iStyleIndex != 0 || iStyleIndexNew != 0)
                                        {
                                            newcell = new SLCell();
                                            newcell.StyleIndex = (uint)iStyleIndex;
                                            newcell.CellText = string.Empty;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                    }
                                }
                                break;
                            case SLPasteTypeValues.Formulas:
                                if (sourcecells.Exists(iRowIndex, iColumnIndex))
                                {
                                    origcell = sourcecells.Cells[iRowIndex][iColumnIndex];
                                    if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                    {
                                        newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                        if (origcell.CellFormula != null) newcell.CellFormula = origcell.CellFormula.Clone();
                                        else newcell.CellFormula = null;
                                        newcell.CellText = origcell.CellText;
                                        newcell.fNumericValue = origcell.fNumericValue;
                                        newcell.DataType = origcell.DataType;
                                        ProcessCellFormulaDelta(ref newcell, false, iStartRowIndex, iStartColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, false, false, 0, 0);
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                    }
                                    else
                                    {
                                        newcell = new SLCell();
                                        if (origcell.CellFormula != null) newcell.CellFormula = origcell.CellFormula.Clone();
                                        else newcell.CellFormula = null;
                                        newcell.CellText = origcell.CellText;
                                        newcell.fNumericValue = origcell.fNumericValue;
                                        newcell.DataType = origcell.DataType;
                                        
                                        iStyleIndexNew = 0;
                                        if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                        if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                        if (iStyleIndexNew != 0) newcell.StyleIndex = (uint)iStyleIndexNew;
                                        ProcessCellFormulaDelta(ref newcell, false, iStartRowIndex, iStartColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, false, false, 0, 0);
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                    }
                                }
                                else
                                {
                                    if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                    {
                                        newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                        newcell.CellText = string.Empty;
                                        newcell.DataType = CellValues.Number;
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                    }
                                    // no else because don't have to do anything
                                }
                                break;
                            case SLPasteTypeValues.Paste:
                                if (sourcecells.Exists(iRowIndex, iColumnIndex))
                                {
                                    origcell = sourcecells.Cells[iRowIndex][iColumnIndex].Clone();
                                    newcell = origcell.Clone();
                                    ProcessCellFormulaDelta(ref newcell, false, iStartRowIndex, iStartColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, false, false, 0, 0);
                                    cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                }
                                else
                                {
                                    // else the source cell is empty
                                    if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                    {
                                        iStyleIndex = 0;
                                        if (sourcerowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)sourcerowstyleindex[iRowIndex];
                                        if (iStyleIndex == 0 && sourcecolstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)sourcecolstyleindex[iColumnIndex];

                                        if (iStyleIndex != 0)
                                        {
                                            newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                            newcell.StyleIndex = (uint)iStyleIndex;
                                            newcell.CellText = string.Empty;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                        else
                                        {
                                            // if the source cell is empty, then direct pasting
                                            // means overwrite the existing cell, which is faster
                                            // by just removing it.
                                            slws.CellWarehouse.Remove(iNewRowIndex, iNewColumnIndex);
                                        }
                                    }
                                    else
                                    {
                                        // else no source and no destination, so we check for row/column
                                        // with non-default styles.
                                        iStyleIndex = 0;
                                        if (sourcerowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)sourcerowstyleindex[iRowIndex];
                                        if (iStyleIndex == 0 && sourcecolstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)sourcecolstyleindex[iColumnIndex];

                                        iStyleIndexNew = 0;
                                        if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                        if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                        if (iStyleIndex != 0 || iStyleIndexNew != 0)
                                        {
                                            newcell = new SLCell();
                                            newcell.StyleIndex = (uint)iStyleIndex;
                                            newcell.CellText = string.Empty;
                                            cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                        }
                                    }
                                }
                                break;
                            case SLPasteTypeValues.Transpose:
                                iNewRowIndex = i - iStartRowIndex;
                                iNewColumnIndex = j - iStartColumnIndex;
                                iSwap = iNewRowIndex;
                                iNewRowIndex = iNewColumnIndex;
                                iNewColumnIndex = iSwap;
                                iNewRowIndex = iNewRowIndex + iStartRowIndex + rowdiff;
                                iNewColumnIndex = iNewColumnIndex + iStartColumnIndex + coldiff;
                                // in case say the millionth row is transposed, because we can't have a millionth column.
                                if (iNewRowIndex <= SLConstants.RowLimit && iNewColumnIndex <= SLConstants.ColumnLimit)
                                {
                                    // this part is identical to normal paste

                                    if (sourcecells.Exists(iRowIndex, iColumnIndex))
                                    {
                                        origcell = sourcecells.Cells[iRowIndex][iColumnIndex].Clone();
                                        ProcessCellFormulaDelta(ref origcell, true, i, j, iNewRowIndex, iNewColumnIndex, false, false, false, 0, 0);
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, origcell);
                                    }
                                    else
                                    {
                                        // else the source cell is empty
                                        if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                        {
                                            iStyleIndex = 0;
                                            if (sourcerowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)sourcerowstyleindex[iRowIndex];
                                            if (iStyleIndex == 0 && sourcecolstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)sourcecolstyleindex[iColumnIndex];

                                            if (iStyleIndex != 0)
                                            {
                                                newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                                newcell.StyleIndex = (uint)iStyleIndex;
                                                newcell.CellText = string.Empty;
                                                cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                            }
                                            else
                                            {
                                                // if the source cell is empty, then direct pasting
                                                // means overwrite the existing cell, which is faster
                                                // by just removing it.
                                                slws.CellWarehouse.Remove(iNewRowIndex, iNewColumnIndex);
                                            }
                                        }
                                        else
                                        {
                                            // else no source and no destination, so we check for row/column
                                            // with non-default styles.
                                            iStyleIndex = 0;
                                            if (sourcerowstyleindex.ContainsKey(iRowIndex)) iStyleIndex = (int)sourcerowstyleindex[iRowIndex];
                                            if (iStyleIndex == 0 && sourcecolstyleindex.ContainsKey(iColumnIndex)) iStyleIndex = (int)sourcecolstyleindex[iColumnIndex];

                                            iStyleIndexNew = 0;
                                            if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                            if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                            if (iStyleIndex != 0 || iStyleIndexNew != 0)
                                            {
                                                newcell = new SLCell();
                                                newcell.StyleIndex = (uint)iStyleIndex;
                                                newcell.CellText = string.Empty;
                                                cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                            }
                                        }
                                    }
                                }
                                break;
                            case SLPasteTypeValues.Values:
                                // this part is identical to the formula part, except
                                // for assigning the cell formula part.

                                if (sourcecells.Exists(iRowIndex, iColumnIndex))
                                {
                                    origcell = sourcecells.Cells[iRowIndex][iColumnIndex];
                                    if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                    {
                                        newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                        newcell.CellFormula = null;
                                        newcell.CellText = origcell.CellText;
                                        newcell.fNumericValue = origcell.fNumericValue;
                                        newcell.DataType = origcell.DataType;
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                    }
                                    else
                                    {
                                        newcell = new SLCell();
                                        newcell.CellFormula = null;
                                        newcell.CellText = origcell.CellText;
                                        newcell.fNumericValue = origcell.fNumericValue;
                                        newcell.DataType = origcell.DataType;

                                        iStyleIndexNew = 0;
                                        if (rowstyleindex.ContainsKey(iNewRowIndex)) iStyleIndexNew = (int)rowstyleindex[iNewRowIndex];
                                        if (iStyleIndexNew == 0 && colstyleindex.ContainsKey(iNewColumnIndex)) iStyleIndexNew = (int)colstyleindex[iNewColumnIndex];

                                        if (iStyleIndexNew != 0) newcell.StyleIndex = (uint)iStyleIndexNew;
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                    }
                                }
                                else
                                {
                                    if (slws.CellWarehouse.Exists(iNewRowIndex, iNewColumnIndex))
                                    {
                                        newcell = slws.CellWarehouse.Cells[iNewRowIndex][iNewColumnIndex].Clone();
                                        newcell.CellFormula = null;
                                        newcell.CellText = string.Empty;
                                        newcell.DataType = CellValues.Number;
                                        cells.SetValue(iNewRowIndex, iNewColumnIndex, newcell);
                                    }
                                    // no else because don't have to do anything
                                }
                                break;
                        }
                    }
                }

                int AnchorEndRowIndex = AnchorRowIndex + iEndRowIndex - iStartRowIndex;
                int AnchorEndColumnIndex = AnchorColumnIndex + iEndColumnIndex - iStartColumnIndex;

                for (i = AnchorRowIndex; i <= AnchorEndRowIndex; ++i)
                {
                    for (j = AnchorColumnIndex; j <= AnchorEndColumnIndex; ++j)
                    {
                        // any cell within destination "paste" operation is taken out
                        slws.CellWarehouse.Remove(i, j);
                    }
                }

                List<int> listRowKeys = cells.Cells.Keys.ToList<int>();
                List<int> listColumnKeys;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = cells.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        origcell = cells.Cells[rowkey][colkey];
                        // the source cells are from another worksheet. Don't know how to rearrange any
                        // cell references in cell formulas...
                        slws.CellWarehouse.SetValue(rowkey, colkey, origcell);
                    }
                }

                // See CopyCell() for the behaviour explanation
                // I'm not going to figure out how to copy merged cells from the source worksheet
                // and decide under what conditions the existing merged cells in the destination
                // worksheet should be removed.
                // So I'm going to just remove any merged cells in the delete range.
                List<SLMergeCell> mca = this.GetWorksheetMergeCells();
                foreach (SLMergeCell mc in mca)
                {
                    if (mc.StartRowIndex >= AnchorRowIndex && mc.EndRowIndex <= AnchorEndRowIndex
                        && mc.StartColumnIndex >= AnchorColumnIndex && mc.EndColumnIndex <= AnchorEndColumnIndex)
                    {
                        slws.MergeCells.Remove(mc);
                    }
                }

                // TODO: conditional formatting and data validations?

                #region Hyperlinks
                if (sourcehyperlinks.Count > 0)
                {
                    // we only care if normal paste or transpose paste. Just like Excel.
                    if (PasteOption == SLPasteTypeValues.Paste || PasteOption == SLPasteTypeValues.Transpose)
                    {
                        List<SLHyperlink> copiedhyperlinks = new List<SLHyperlink>();
                        SLHyperlink hlCopied;

                        int iOverlapStartRowIndex = 1;
                        int iOverlapStartColumnIndex = 1;
                        int iOverlapEndRowIndex = 1;
                        int iOverlapEndColumnIndex = 1;
                        foreach (SLHyperlink hl in sourcehyperlinks)
                        {
                            // this comes from the separating axis theorem.
                            // See merged cells for more details.
                            // In this case however, we're doing stuff when there's overlapping.
                            if (!(iEndRowIndex < hl.Reference.StartRowIndex
                                || iStartRowIndex > hl.Reference.EndRowIndex
                                || iEndColumnIndex < hl.Reference.StartColumnIndex
                                || iStartColumnIndex > hl.Reference.EndColumnIndex))
                            {
                                // get the overlapping region
                                iOverlapStartRowIndex = Math.Max(iStartRowIndex, hl.Reference.StartRowIndex);
                                iOverlapStartColumnIndex = Math.Max(iStartColumnIndex, hl.Reference.StartColumnIndex);
                                iOverlapEndRowIndex = Math.Min(iEndRowIndex, hl.Reference.EndRowIndex);
                                iOverlapEndColumnIndex = Math.Min(iEndColumnIndex, hl.Reference.EndColumnIndex);

                                // offset to the correctly pasted region
                                if (PasteOption == SLPasteTypeValues.Paste)
                                {
                                    iOverlapStartRowIndex += rowdiff;
                                    iOverlapStartColumnIndex += coldiff;
                                    iOverlapEndRowIndex += rowdiff;
                                    iOverlapEndColumnIndex += coldiff;
                                }
                                else
                                {
                                    // can only be transpose. See if check above.

                                    if (iOverlapEndRowIndex > SLConstants.ColumnLimit)
                                    {
                                        // probably won't happen. This means that after transpose,
                                        // the end row index will flip to exceed the column limit.
                                        // I don't feel like testing how Excel handles this, so
                                        // I'm going to just take it as normal paste.
                                        iOverlapStartRowIndex += rowdiff;
                                        iOverlapStartColumnIndex += coldiff;
                                        iOverlapEndRowIndex += rowdiff;
                                        iOverlapEndColumnIndex += coldiff;
                                    }
                                    else
                                    {
                                        iOverlapStartRowIndex -= iStartRowIndex;
                                        iOverlapStartColumnIndex -= iStartColumnIndex;
                                        iOverlapEndRowIndex -= iStartRowIndex;
                                        iOverlapEndColumnIndex -= iStartColumnIndex;

                                        iSwap = iOverlapStartRowIndex;
                                        iOverlapStartRowIndex = iOverlapStartColumnIndex;
                                        iOverlapStartColumnIndex = iSwap;

                                        iSwap = iOverlapEndRowIndex;
                                        iOverlapEndRowIndex = iOverlapEndColumnIndex;
                                        iOverlapEndColumnIndex = iSwap;

                                        iOverlapStartRowIndex += (iStartRowIndex + rowdiff);
                                        iOverlapStartColumnIndex += (iStartColumnIndex + coldiff);
                                        iOverlapEndRowIndex += (iStartRowIndex + rowdiff);
                                        iOverlapEndColumnIndex += (iStartColumnIndex + coldiff);
                                    }
                                }

                                hlCopied = new SLHyperlink();
                                hlCopied = hl.Clone();
                                hlCopied.IsNew = true;
                                if (hlCopied.IsExternal)
                                {
                                    if (hlurl.ContainsKey(hlCopied.Id))
                                    {
                                        hlCopied.HyperlinkUri = hlurl[hlCopied.Id];
                                        if (hlCopied.HyperlinkUri.StartsWith("."))
                                        {
                                            // assume this is a relative file path such as ../ or ./
                                            hlCopied.HyperlinkUriKind = UriKind.Relative;
                                        }
                                        else
                                        {
                                            hlCopied.HyperlinkUriKind = UriKind.Absolute;
                                        }
                                        hlCopied.Id = string.Empty;
                                    }
                                }
                                hlCopied.Reference = new SLCellPointRange(iOverlapStartRowIndex, iOverlapStartColumnIndex, iOverlapEndRowIndex, iOverlapEndColumnIndex);
                                copiedhyperlinks.Add(hlCopied);
                            }
                        }

                        if (copiedhyperlinks.Count > 0)
                        {
                            slws.Hyperlinks.AddRange(copiedhyperlinks);
                        }
                    }
                }
                #endregion

                #region Calculation cells
                if (slwb.CalculationCells.Count > 0)
                {
                    List<int> listToDelete = new List<int>();
                    for (i = 0; i < slwb.CalculationCells.Count; ++i)
                    {
                        if (slwb.CalculationCells[i].SheetId == giSelectedWorksheetID)
                        {
                            iRowIndex = slwb.CalculationCells[i].RowIndex;
                            iColumnIndex = slwb.CalculationCells[i].ColumnIndex;

                            if (iRowIndex >= AnchorRowIndex && iRowIndex <= AnchorEndRowIndex
                                && iColumnIndex >= AnchorColumnIndex && iColumnIndex <= AnchorEndColumnIndex)
                            {
                                // existing calculation cell lies within destination "paste" operation
                                if (!listToDelete.Contains(i)) listToDelete.Add(i);
                            }
                        }
                    }

                    for (i = listToDelete.Count - 1; i >= 0; --i)
                    {
                        slwb.CalculationCells.RemoveAt(listToDelete[i]);
                    }
                }
                #endregion
            }

            return result;
        }

        /// <summary>
        /// Clear all cell content in the worksheet.
        /// </summary>
        /// <returns>True if content has been cleared. False otherwise. If there are no content in the worksheet, false is also returned.</returns>
        public bool ClearCellContent()
        {
            bool result = false;
            List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
            List<int> listColumnKeys;
            foreach (int rowkey in listRowKeys)
            {
                listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                foreach (int colkey in listColumnKeys)
                {
                    this.ClearCellContentData(rowkey, colkey);
                }
            }

            return result;
        }

        /// <summary>
        /// Clear all cell content within specified rows and columns. If the top-left cell of a merged cell is within specified rows and columns, the merged cell content is also cleared.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be cleared, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be cleared, such as "A1". This is typically the bottom-right cell.</param>
        /// <returns>True if content has been cleared. False otherwise. If there are no content within specified rows and columns, false is also returned.</returns>
        public bool ClearCellContent(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex)
                || !SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                return false;
            }

            return ClearCellContent(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Clear all cell content within specified rows and columns. If the top-left cell of a merged cell is within specified rows and columns, the merged cell content is also cleared.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <returns>True if content has been cleared. False otherwise. If there are no content within specified rows and columns, false is also returned.</returns>
        public bool ClearCellContent(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            bool result = false;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            long iSize = (iEndRowIndex - iStartRowIndex + 1) * (iEndColumnIndex - iStartColumnIndex + 1);

            int iRowIndex = -1, iColumnIndex = -1;
            for (iRowIndex = iStartRowIndex; iRowIndex <= iEndRowIndex; ++iRowIndex)
            {
                for (iColumnIndex = iStartColumnIndex; iColumnIndex <= iEndColumnIndex; ++iColumnIndex)
                {
                    if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                    {
                        this.ClearCellContentData(iRowIndex, iColumnIndex);
                        result = true;
                    }
                }
            }

            List<int> listToDelete = new List<int>();
            int i;
            for (i = 0; i < slwb.CalculationCells.Count; ++i)
            {
                if (slwb.CalculationCells[i].SheetId == giSelectedWorksheetID)
                {
                    iRowIndex = slwb.CalculationCells[i].RowIndex;
                    iColumnIndex = slwb.CalculationCells[i].ColumnIndex;
                    if (iRowIndex >= iStartRowIndex && iRowIndex <= iEndRowIndex
                        && iColumnIndex >= iStartColumnIndex && iColumnIndex <= iEndColumnIndex)
                    {
                        if (!listToDelete.Contains(i)) listToDelete.Add(i);
                    }
                }
            }

            for (i = listToDelete.Count - 1; i >= 0; --i)
            {
                slwb.CalculationCells.RemoveAt(listToDelete[i]);
            }

            return result;
        }

        private void ClearCellContentData(int RowIndex, int ColumnIndex)
        {
            if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
            {
                SLCell c = slws.CellWarehouse.Cells[RowIndex][ColumnIndex].Clone();
                c.CellFormula = null;
                c.DataType = CellValues.Number;
                c.NumericValue = 0;
                // if the cell still has attributes (say the style index), then update it
                // otherwise remove the cell
                if (c.StyleIndex != 0 || c.CellMetaIndex != 0 || c.ValueMetaIndex != 0 || c.ShowPhonetic != false)
                {
                    slws.CellWarehouse.SetValue(RowIndex, ColumnIndex, c);
                }
                else
                {
                    slws.CellWarehouse.Remove(RowIndex, ColumnIndex);
                }
            }
        }

        /// <summary>
        /// Get existing shared cell formulas in the current worksheet in a list of SLSharedCellFormula objects.
        /// NOTE: Due to technical difficulties (read: a certain popular spreadsheet software's behaviour is confusing),
        /// any copy/insert/delete of cells/rows/columns will flatten all shared cell formula into the respective cells.
        /// WARNING: This is only a snapshot. Any changes made to the returned result are not used.
        /// </summary>
        /// <returns>A list of existing shared cell formulas.</returns>
        public List<SLSharedCellFormula> GetSharedCellFormulas()
        {
            List<SLSharedCellFormula> result = new List<SLSharedCellFormula>();
            List<uint> keys = slws.SharedCellFormulas.Keys.ToList<uint>();
            keys.Sort();
            for (int i = 0; i < keys.Count; ++i)
            {
                result.Add(slws.SharedCellFormulas[keys[i]].Clone());
            }

            return result;
        }

        /// <summary>
        /// Get the cell formula if it exists.
        /// </summary>
        /// <param name="CellReference">The cell reference, such as "A1".</param>
        /// <returns>The cell formula.</returns>
        public string GetCellFormula(string CellReference)
        {
            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                return string.Empty;
            }

            return GetCellFormula(iRowIndex, iColumnIndex);
        }

        /// <summary>
        /// Get the cell formula if it exists.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The cell formula.</returns>
        public string GetCellFormula(int RowIndex, int ColumnIndex)
        {
            string result = string.Empty;
            if (RowIndex >= 1 && RowIndex <= SLConstants.RowLimit && ColumnIndex >= 1 && ColumnIndex <= SLConstants.ColumnLimit)
            {
                // check only if there's an existing cell
                if (slws.CellWarehouse.Exists(RowIndex, ColumnIndex))
                {
                    // There are 4 types of formulas: Array, DataTable, Normal and Shared
                    // Normal is when the formula is embedded directly with the Cell class.
                    // Shared is when the formula is shared (duh), subject to Open XML specs and rules.
                    // I don't know what the other 2 types do...

                    bool bHasError;
                    SLSharedCellFormula scf;
                    int i, j;
                    bool bFound = false;
                    List<uint> list = slws.SharedCellFormulas.Keys.ToList<uint>();
                    for (i = 0; i < list.Count; ++i)
                    {
                        scf = slws.SharedCellFormulas[list[i]];
                        for (j = 0; j < scf.Reference.Count; ++j)
                        {
                            // if within reference bounds
                            if (scf.Reference[j].StartRowIndex <= RowIndex && RowIndex <= scf.Reference[j].EndRowIndex
                                && scf.Reference[j].StartColumnIndex <= ColumnIndex && ColumnIndex <= scf.Reference[j].EndColumnIndex)
                            {
                                bHasError = false;
                                result = AdjustCellFormulaDelta(scf.FormulaText, false, scf.BaseCellRowIndex, scf.BaseCellColumnIndex, RowIndex, ColumnIndex, false, false, false, false, 0, 0, out bHasError);
                                bFound = true;
                                break;
                            }
                        }

                        if (bFound) break;
                    }

                    SLCell cell = slws.CellWarehouse.Cells[RowIndex][ColumnIndex];
                    if (!bFound && cell.CellFormula != null)
                    {
                        result = cell.CellFormula.FormulaText;
                    }
                }
            }

            return result;
        }

        internal void SwapRangeIndexIfNecessary(ref string SheetName1, ref string SheetName2, ref int Index1, ref int Index2, ref bool IsAbsolute1, ref bool IsAbsolute2)
        {
            if (SheetName1.Equals(SheetName2, StringComparison.InvariantCultureIgnoreCase) && Index1 > Index2)
            {
                string sSwap;
                int iSwap;
                bool bSwap;

                sSwap = SheetName1;
                SheetName1 = SheetName2;
                SheetName2 = sSwap;

                iSwap = Index1;
                Index1 = Index2;
                Index2 = iSwap;

                bSwap = IsAbsolute1;
                IsAbsolute1 = IsAbsolute2;
                IsAbsolute2 = bSwap;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="ToSwapRowColumn">this is for copying on transpose</param>
        /// <param name="StartRowIndex"></param>
        /// <param name="StartColumnIndex"></param>
        /// <param name="AnchorRowIndex"></param>
        /// <param name="AnchorColumnIndex"></param>
        /// <param name="CheckForInsertDeleteRowColumn">true when it's insert/delete row column operation. otherwise it's copy cell/row/column</param>
        /// <param name="CheckForInsert">true when it's insert row/column, false when it's delete row/column</param>
        /// <param name="CheckForRow"></param>
        /// <param name="CheckStartIndex">only used when delete. put 0 when insert.</param>
        /// <param name="CheckEndIndex">only used when delete. put 0 when insert.</param>
        internal void ProcessCellFormulaDelta(ref SLCell cell, bool ToSwapRowColumn, int StartRowIndex, int StartColumnIndex, int AnchorRowIndex, int AnchorColumnIndex, bool CheckForInsertDeleteRowColumn, bool CheckForInsert, bool CheckForRow, int CheckStartIndex, int CheckEndIndex)
        {
            // don't have to do anything if there's no change to rows or columns
            if (CheckForInsertDeleteRowColumn || StartRowIndex != AnchorRowIndex || StartColumnIndex != AnchorColumnIndex)
            {
                bool bHasError = false;
                if (cell.CellText != null && cell.CellText.StartsWith("="))
                {
                    //...cell.CellText = AddDeleteCellFormulaDelta(cell.CellText, StartRowIndex, RowDelta, StartColumnIndex, ColumnDelta);
                    cell.CellText = AdjustCellFormulaDelta(cell.CellText, ToSwapRowColumn, StartRowIndex, StartColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, CheckForInsertDeleteRowColumn, CheckForInsert, CheckForRow, CheckStartIndex, CheckEndIndex, out bHasError);
                    if (bHasError)
                    {
                        cell.DataType = CellValues.Error;
                        cell.CellText = SLConstants.ErrorReference;
                    }
                }
                if (cell.CellFormula != null)
                {
                    if (cell.CellFormula.FormulaType == CellFormulaValues.Normal)
                    {
                        //...cell.CellFormula.FormulaText = AddDeleteCellFormulaDelta(cell.CellFormula.FormulaText, StartRowIndex, RowDelta, StartColumnIndex, ColumnDelta);
                        cell.CellFormula.FormulaText = AdjustCellFormulaDelta(cell.CellFormula.FormulaText, ToSwapRowColumn, StartRowIndex, StartColumnIndex, AnchorRowIndex, AnchorColumnIndex, false, CheckForInsertDeleteRowColumn, CheckForInsert, CheckForRow, CheckStartIndex, CheckEndIndex, out bHasError);
                        if (bHasError)
                        {
                            cell.DataType = CellValues.Error;
                            cell.CellText = SLConstants.ErrorReference;
                        }
                        else
                        {
                            // because we don't know how to calculate formulas yet
                            cell.CellText = string.Empty;
                        }
                    }
                }
            }
        }

        //CellFormula.SharedIndex see for details about shared formula and si

        /// <summary>
        /// 
        /// </summary>
        /// <param name="CellFormula"></param>
        /// <param name="ToSwapRowColumn">this is for copying on transpose</param>
        /// <param name="StartRowIndex"></param>
        /// <param name="StartColumnIndex"></param>
        /// <param name="AnchorRowIndex"></param>
        /// <param name="AnchorColumnIndex"></param>
        /// <param name="OnlyForCurrentWorksheet"></param>
        /// <param name="CheckForInsertDeleteRowColumn">true when it's insert/delete row column operation. otherwise it's copy cell/row/column</param>
        /// <param name="CheckForInsert">true when it's insert row/column, false when it's delete row/column</param>
        /// <param name="CheckForRow"></param>
        /// <param name="CheckStartIndex">only used when delete. put 0 when insert.</param>
        /// <param name="CheckEndIndex">only used when delete. put 0 when insert.</param>
        /// <param name="HasError"></param>
        /// <returns></returns>
        internal string AdjustCellFormulaDelta(string CellFormula, bool ToSwapRowColumn, int StartRowIndex, int StartColumnIndex, int AnchorRowIndex, int AnchorColumnIndex, bool OnlyForCurrentWorksheet, bool CheckForInsertDeleteRowColumn, bool CheckForInsert, bool CheckForRow, int CheckStartIndex, int CheckEndIndex, out bool HasError)
        {
            HasError = false;

            // if there's no change, just return
            if (!CheckForInsertDeleteRowColumn && StartRowIndex == AnchorRowIndex && StartColumnIndex == AnchorColumnIndex) return CellFormula;

            // The ToSwapRowColumn parameter is for transposing cell references in formulas when copying.
            // When transposing, you move the cell indices to the origin (top-left corner of the spreadsheet in this case),
            // then swap the row and column indices (you know, because of the transposing), then move the resulting
            // cell indices to the anchor cell.
            // For example, let's say the cell E5 has the formula "=E1+G2". And we copy transpose-wise to cell G14.
            // E5 means start row index is 5 and start column index is 5 (because E).
            // G14 means anchor row index is 14 and anchor column index is 7 (because G).
            // So "E1" (from the formula) is row 1 column 5.
            // Moving that to the origin means row = 1 - 5(start row), column = 5 - 5(start column)
            // ==> row = -4, column = 0
            // Yes, the row index is negative. Don't panic.
            // Swapping the indices, we get row = 0, column = -4
            // Then move to the anchor, we get row = 0 + 14, column = -4 + 7
            // ==> row = 14, column = 3
            // Which is the cell C14.
            // Let's do it one more time for "G2" (from the formula).
            // "G2" is row 2 column 7.
            // Moving to origin means row = 2 - 5(start row), column = 7 - 5(start column)
            // ==> row = -3, column = 2
            // Swapping, we get row = 2, column = -3
            // Then move to the anchor, we get row = 2 + 14, column -3 + 7
            // ==> row = 16, column = 4
            // Which becomes the cell D16.

            // If you do any OpenGL (or DirectX) programming, the above concept will be familiar.
            // To rotate a 3D object about its own centre in space, you first transpose it to the origin,
            // then rotate it, then transpose it back to its original position.
            // In this case, we don't transpose it back to the original position (start cell reference)
            // but to the anchor position instead.
            // SIDE NOTE: the term "transpose" in 3D programming means "move". The term "transpose" in Excel
            // means "rows become columns and columns become rows" (also in maths, as in matrix transpose).

            // The assumption is that the formula is well-formed and that literal strings
            // have matching quotes (meaning there's always a start and end double quote, or not at all)
            // Well-formed-ness includes the criteria that functions are always immediately followed
            // by a starting round bracket (.
            // Well-formed-ness means generally anything Excel will accept as valid formula.
            // For example "=SUM(A1:B3)" is well-formed, but "=SUM (A1:B3)" is not.
            // Or at least last I checked, Excel requires the formula functions to be like that.
            // Row-only or column-only formulas seem to only have the sheet name on the first index if at all.
            // For example, "=SUM(Sheet1!3:5)" or "=SUM(3:5)".
            // There doesn't seem to be "=SUM(Sheet1!3:Sheet1!5)" [Excel thinks this is an error]
            // I'm going to take note of the "second" worksheet name in any case...
            List<string> listFormula = new List<string>();
            // only #REF! or take care of the other error types as well?
            string sErrorReferenceMatch = string.Format("((\\w+[!])?{0})|((\\w+[!])?{1})|((\\w+[!])?{2})|((\\w+[!])?{3})|((\\w+[!])?{4})|((\\w+[!])?{5})|((\\w+[!])?{6})", SLConstants.ErrorDivisionByZero, SLConstants.ErrorNA, SLConstants.ErrorName, SLConstants.ErrorNull, SLConstants.ErrorNumber, SLConstants.ErrorReference, SLConstants.ErrorValue);
            string sCellReferenceMatch = "(\\w+[!])?[$]?[a-zA-Z]{1,3}[$]?\\d{1,7}";
            string sRowReferenceMatch = "(\\w+[!])?[$]?\\d{1,7}";
            string sColumnReferenceMatch = "(\\w+[!])?[$]?[a-zA-Z]{1,3}";
            // match errors first, then cell references, then rows then columns
            // cell references need the [(]? because of LOG10 as cell reference
            // and LOG10(number) as a function.
            // This requires that formulas are well-formed (see above comment).
            // Maybe Excel uses this to differentiate between the two LOG10 too.
            // Row/column-only references must(???) always come in a pair in cell formulas
            // Cell references can be single or in a range, hence the (:{0})? part
            Regex rgx = new Regex(string.Format("({0})|({1}(:{1})?[(]?)|({2}:{2})|({3}:{3})", sErrorReferenceMatch, sCellReferenceMatch, sRowReferenceMatch, sColumnReferenceMatch));
            Regex rgxError = new Regex(string.Format("^{0}$", sErrorReferenceMatch));
            Regex rgxCell = new Regex(string.Format("^{0}(:{0})?$", sCellReferenceMatch));
            Regex rgxRow = new Regex(string.Format("^{0}:{0}$", sRowReferenceMatch));
            Regex rgxColumn = new Regex(string.Format("^{0}:{0}$", sColumnReferenceMatch));

            // to split literal strings such as =CONCATENATE("B1", "C1:D3", B1) into
            // index 0:=CONCATENATE(
            // index 1:B1
            // index 2:, ((note the space after the comma))
            // index 3:C1:D3
            // index 4:, B1)
            // The odd numbered entries are thus literal strings, and so we don't do matching for them.
            string[] saLiteralStringSplit = CellFormula.Split("\"".ToCharArray(), StringSplitOptions.None);

            // this is used for optimisation purpose. Most of the time, you don't swap,
            // so just use this value, which is faster than calculating the delta every time.
            int iRowDelta = AnchorRowIndex - StartRowIndex;
            int iColumnDelta = AnchorColumnIndex - StartColumnIndex;

            Match m;
            int i;
            string sMatch;
            for (i = 0; i < saLiteralStringSplit.Length; ++i)
            {
                if (i % 2 == 0)
                {
                    while (rgx.IsMatch(saLiteralStringSplit[i]))
                    {
                        m = rgx.Match(saLiteralStringSplit[i]);
                        sMatch = m.ToString();
                        listFormula.Add(saLiteralStringSplit[i].Substring(0, m.Index));
                        listFormula.Add(sMatch);
                        saLiteralStringSplit[i] = saLiteralStringSplit[i].Substring(sMatch.Length + m.Index);
                    }
                    listFormula.Add(saLiteralStringSplit[i]);
                }
                else
                {
                    // add the literal string. Need to put the double quotes back!
                    listFormula.Add(string.Format("\"{0}\"", saLiteralStringSplit[i]));
                }
            }

            string sCellRef1 = string.Empty, sCellRef2 = string.Empty;
            string sSheetName1 = string.Empty, sSheetName2 = string.Empty;
            bool bIsCurrentSheet = false;
            string sFormula = string.Empty;
            int iRowIndex1 = 0, iRowIndex2 = 0;
            int iColumnIndex1 = 0, iColumnIndex2 = 0;
            int index = 0, index2 = 0;
            int iSwap = 0;
            bool bSwap = false;
            bool bIsAbsoluteRow1 = false, bIsAbsoluteRow2 = false;
            bool bIsAbsoluteColumn1 = false, bIsAbsoluteColumn2 = false;

            bool bHasCheckError = false;
            for (i = 0; i < listFormula.Count; ++i)
            {
                if (rgxError.IsMatch(listFormula[i]))
                {
                    // leave the original error message alone
                    HasError = true;
                }
                else if (rgxCell.IsMatch(listFormula[i]))
                {
                    #region for cell
                    index = listFormula[i].IndexOf(":");
                    if (index > -1)
                    {
                        sCellRef1 = listFormula[i].Substring(0, index);
                        sCellRef2 = listFormula[i].Substring(index + 1);

                        sSheetName1 = string.Empty;
                        index2 = sCellRef1.IndexOf("!");
                        if (index2 > -1)
                        {
                            sSheetName1 = sCellRef1.Substring(0, index2);
                            sCellRef1 = sCellRef1.Substring(index2 + 1);
                        }

                        sSheetName2 = string.Empty;
                        index2 = sCellRef2.IndexOf("!");
                        if (index2 > -1)
                        {
                            sSheetName2 = sCellRef2.Substring(0, index2);
                            sCellRef2 = sCellRef2.Substring(index2 + 1);
                        }

                        bIsAbsoluteRow1 = false;
                        bIsAbsoluteRow2 = false;
                        bIsAbsoluteColumn1 = false;
                        bIsAbsoluteColumn2 = false;

                        if (Regex.IsMatch(sCellRef1, "\\$[a-zA-Z]{1,3}"))
                        {
                            bIsAbsoluteColumn1 = true;
                        }
                        if (Regex.IsMatch(sCellRef1, "\\$\\d{1,7}"))
                        {
                            bIsAbsoluteRow1 = true;
                        }

                        if (Regex.IsMatch(sCellRef2, "\\$[a-zA-Z]{1,3}"))
                        {
                            bIsAbsoluteColumn2 = true;
                        }
                        if (Regex.IsMatch(sCellRef2, "\\$\\d{1,7}"))
                        {
                            bIsAbsoluteRow2 = true;
                        }

                        // we remove the dollar signs
                        if (bIsAbsoluteRow1 || bIsAbsoluteColumn1) sCellRef1 = sCellRef1.Replace("$", "");
                        if (bIsAbsoluteRow2 || bIsAbsoluteColumn2) sCellRef2 = sCellRef2.Replace("$", "");

                        if (SLTool.FormatCellReferenceToRowColumnIndex(sCellRef1, out iRowIndex1, out iColumnIndex1)
                            && SLTool.FormatCellReferenceToRowColumnIndex(sCellRef2, out iRowIndex2, out iColumnIndex2))
                        {
                            // Excel seems to swap the indices of ranges to fit a top-left/bottom-right
                            // range, BUT only if the sheet names are empty.
                            // For example, "=SUM(E1:A7)" becomes "=SUM(A1:E7)"
                            // This wouldn't happen if you're using Excel because Excel would already
                            // have corrected you when you're entering the formula.
                            // But hey, I'm trying to be helpful and imitating Excel behaviour.
                            // Sometimes, I think I do too much...
                            if (iRowIndex1 > iRowIndex2 && sSheetName1.Length == 0 && sSheetName2.Length == 0)
                            {
                                iSwap = iRowIndex1;
                                iRowIndex1 = iRowIndex2;
                                iRowIndex2 = iSwap;

                                bSwap = bIsAbsoluteRow1;
                                bIsAbsoluteRow1 = bIsAbsoluteRow2;
                                bIsAbsoluteRow2 = bSwap;
                            }

                            if (iColumnIndex1 > iColumnIndex2 && sSheetName1.Length == 0 && sSheetName2.Length == 0)
                            {
                                iSwap = iColumnIndex1;
                                iColumnIndex1 = iColumnIndex2;
                                iColumnIndex2 = iSwap;

                                bSwap = bIsAbsoluteColumn1;
                                bIsAbsoluteColumn1 = bIsAbsoluteColumn2;
                                bIsAbsoluteColumn2 = bSwap;
                            }

                            // Pseudo code logic:
                            // if abs, don't change
                            // if not abs,
                            //     if onlyforcurrent
                            //         if (gsCurrent.equals(sheetname))
                            //             change
                            //         else
                            //             no change
                            //     else
                            //         change

                            // But there are complications, because I don't know the full range of valid
                            // cell references. For example, did you know this
                            // SUM(Sheet1!A1:Sheet1!B3)
                            // is a valid formula? But it's also SUM(A1:B3) if the current sheet is Sheet1.
                            // Also SUM(Sheet1!A1:B3) is the same.
                            // But at least Excel rejects this SUM(Sheet1!A1:Sheet2!B3)

                            // So we have the following additional checks for current name:
                            // is current when sheetname1 and sheetname2 are empty string
                            // or is current when sheetname1 is currentname and sheetname2 is empty string
                            // or is current when sheetname1 is empty string and sheetname2 is currentname
                            // or is current when sheetname1 is currentname and sheetname2 is currentname

                            bIsCurrentSheet = (string.IsNullOrEmpty(sSheetName1) && string.IsNullOrEmpty(sSheetName2));
                            bIsCurrentSheet |= (gsSelectedWorksheetName.Equals(sSheetName1, StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(sSheetName2));
                            bIsCurrentSheet |= (string.IsNullOrEmpty(sSheetName1) && gsSelectedWorksheetName.Equals(sSheetName2, StringComparison.OrdinalIgnoreCase));
                            bIsCurrentSheet |= (gsSelectedWorksheetName.Equals(sSheetName1, StringComparison.OrdinalIgnoreCase) && gsSelectedWorksheetName.Equals(sSheetName2, StringComparison.OrdinalIgnoreCase));

                            if (ToSwapRowColumn)
                            {
                                // no need for the Check* related variables because swapping will not involve the checks.

                                // if any absolutes, no swap
                                if (!bIsAbsoluteRow1 && !bIsAbsoluteColumn1)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        iRowIndex1 -= StartRowIndex;
                                        iColumnIndex1 -= StartColumnIndex;

                                        iSwap = iRowIndex1;
                                        iRowIndex1 = iColumnIndex1;
                                        iColumnIndex1 = iSwap;

                                        iRowIndex1 += AnchorRowIndex;
                                        iColumnIndex1 += AnchorColumnIndex;
                                    }
                                }

                                // if any absolutes, no swap
                                if (!bIsAbsoluteRow2 && !bIsAbsoluteColumn2)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        iRowIndex2 -= StartRowIndex;
                                        iColumnIndex2 -= StartColumnIndex;

                                        iSwap = iRowIndex2;
                                        iRowIndex2 = iColumnIndex2;
                                        iColumnIndex2 = iSwap;

                                        iRowIndex2 += AnchorRowIndex;
                                        iColumnIndex2 += AnchorColumnIndex;
                                    }
                                }
                            }
                            else
                            {
                                bHasCheckError = false;
                                // check for delete range
                                if (CheckForInsertDeleteRowColumn && !CheckForInsert)
                                {
                                    if (CheckForRow)
                                    {
                                        if (CheckStartIndex <= iRowIndex1 && iRowIndex2 <= CheckEndIndex)
                                        {
                                            bHasCheckError = true;
                                        }
                                    }
                                    else
                                    {
                                        if (CheckStartIndex <= iColumnIndex1 && iColumnIndex2 <= CheckEndIndex)
                                        {
                                            bHasCheckError = true;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteRow1)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && CheckForRow && iRowDelta > 0)
                                        {
                                            if (iRowIndex1 >= StartRowIndex) iRowIndex1 += iRowDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && CheckForRow && iRowDelta < 0)
                                        {
                                            if (iRowIndex1 >= StartRowIndex) iRowIndex1 += iRowDelta;
                                        }
                                        else
                                        {
                                            iRowIndex1 += iRowDelta;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteRow2)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && CheckForRow && iRowDelta > 0)
                                        {
                                            if (iRowIndex2 >= StartRowIndex) iRowIndex2 += iRowDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && CheckForRow && iRowDelta < 0)
                                        {
                                            if (iRowIndex2 >= StartRowIndex) iRowIndex2 += iRowDelta;
                                        }
                                        else
                                        {
                                            iRowIndex2 += iRowDelta;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteColumn1)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && !CheckForRow && iColumnDelta > 0)
                                        {
                                            if (iColumnIndex1 >= StartColumnIndex) iColumnIndex1 += iColumnDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && !CheckForRow && iColumnDelta < 0)
                                        {
                                            if (iColumnIndex1 >= StartColumnIndex) iColumnIndex1 += iColumnDelta;
                                        }
                                        else
                                        {
                                            iColumnIndex1 += iColumnDelta;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteColumn2)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && !CheckForRow && iColumnDelta > 0)
                                        {
                                            if (iColumnIndex2 >= StartColumnIndex) iColumnIndex2 += iColumnDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && !CheckForRow && iColumnDelta < 0)
                                        {
                                            if (iColumnIndex2 >= StartColumnIndex) iColumnIndex2 += iColumnDelta;
                                        }
                                        else
                                        {
                                            iColumnIndex2 += iColumnDelta;
                                        }
                                    }
                                }
                            }

                            // override the delta changes above if there's an error.
                            if (bHasCheckError)
                            {
                                iRowIndex1 = -1;
                                iRowIndex2 = -1;
                                iColumnIndex1 = -1;
                                iColumnIndex2 = -1;
                            }

                            if (iRowIndex1 < 1 || iRowIndex1 > SLConstants.RowLimit
                                || iColumnIndex1 < 1 || iColumnIndex1 > SLConstants.ColumnLimit
                                || iRowIndex2 < 1 || iRowIndex2 > SLConstants.RowLimit
                                || iColumnIndex2 < 1 || iColumnIndex2 > SLConstants.ColumnLimit)
                            {
                                listFormula[i] = string.Format("{0}{1}", sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty, SLConstants.ErrorReference);
                                HasError = true;
                            }
                            else
                            {
                                listFormula[i] = string.Format("{0}{1}{2}:{3}{4}{5}",
                                    sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty,
                                    bIsAbsoluteColumn1 ? "$" + SLTool.ToColumnName(iColumnIndex1) : SLTool.ToColumnName(iColumnIndex1),
                                    bIsAbsoluteRow1 ? "$" + iRowIndex1.ToString(CultureInfo.InvariantCulture) : iRowIndex1.ToString(CultureInfo.InvariantCulture),
                                    sSheetName2.Length > 0 ? sSheetName2 + "!" : string.Empty,
                                    bIsAbsoluteColumn2 ? "$" + SLTool.ToColumnName(iColumnIndex2) : SLTool.ToColumnName(iColumnIndex2),
                                    bIsAbsoluteRow2 ? "$" + iRowIndex2.ToString(CultureInfo.InvariantCulture) : iRowIndex2.ToString(CultureInfo.InvariantCulture));
                            }
                        }
                    }
                    else
                    {
                        // else is a single cell reference
                        sCellRef1 = listFormula[i];

                        sSheetName1 = string.Empty;
                        index2 = sCellRef1.IndexOf("!");
                        if (index2 > -1)
                        {
                            sSheetName1 = sCellRef1.Substring(0, index2);
                            sCellRef1 = sCellRef1.Substring(index2 + 1);
                        }

                        bIsAbsoluteRow1 = false;
                        bIsAbsoluteColumn1 = false;

                        if (Regex.IsMatch(sCellRef1, "\\$[a-zA-Z]{1,3}"))
                        {
                            bIsAbsoluteColumn1 = true;
                        }
                        if (Regex.IsMatch(sCellRef1, "\\$\\d{1,7}"))
                        {
                            bIsAbsoluteRow1 = true;
                        }

                        // we remove the dollar signs
                        if (bIsAbsoluteRow1 || bIsAbsoluteColumn1) sCellRef1 = sCellRef1.Replace("$", "");

                        if (SLTool.FormatCellReferenceToRowColumnIndex(sCellRef1, out iRowIndex1, out iColumnIndex1))
                        {
                            bIsCurrentSheet = string.IsNullOrEmpty(sSheetName1) || gsSelectedWorksheetName.Equals(sSheetName1, StringComparison.OrdinalIgnoreCase);

                            if (ToSwapRowColumn)
                            {
                                // if any absolutes, no swap
                                if (!bIsAbsoluteRow1 && !bIsAbsoluteColumn1)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        iRowIndex1 -= StartRowIndex;
                                        iColumnIndex1 -= StartColumnIndex;

                                        iSwap = iRowIndex1;
                                        iRowIndex1 = iColumnIndex1;
                                        iColumnIndex1 = iSwap;

                                        iRowIndex1 += AnchorRowIndex;
                                        iColumnIndex1 += AnchorColumnIndex;
                                    }
                                }
                            }
                            else
                            {
                                bHasCheckError = false;
                                // check for delete range
                                if (CheckForInsertDeleteRowColumn && !CheckForInsert)
                                {
                                    if (CheckForRow)
                                    {
                                        if (CheckStartIndex <= iRowIndex1 && iRowIndex1 <= CheckEndIndex)
                                        {
                                            bHasCheckError = true;
                                        }
                                    }
                                    else
                                    {
                                        if (CheckStartIndex <= iColumnIndex1 && iColumnIndex1 <= CheckEndIndex)
                                        {
                                            bHasCheckError = true;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteRow1)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && CheckForRow && iRowDelta > 0)
                                        {
                                            if (iRowIndex1 >= StartRowIndex) iRowIndex1 += iRowDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && CheckForRow && iRowDelta < 0)
                                        {
                                            if (iRowIndex1 >= StartRowIndex) iRowIndex1 += iRowDelta;
                                        }
                                        else
                                        {
                                            iRowIndex1 += iRowDelta;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteColumn1)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && !CheckForRow && iColumnDelta > 0)
                                        {
                                            if (iColumnIndex1 >= StartColumnIndex) iColumnIndex1 += iColumnDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && !CheckForRow && iColumnDelta < 0)
                                        {
                                            if (iColumnIndex1 >= StartColumnIndex) iColumnIndex1 += iColumnDelta;
                                        }
                                        else
                                        {
                                            iColumnIndex1 += iColumnDelta;
                                        }
                                    }
                                }

                                // override the delta changes above if there's an error.
                                if (bHasCheckError)
                                {
                                    iRowIndex1 = -1;
                                    iColumnIndex1 = -1;
                                }
                            }

                            if (iRowIndex1 < 1 || iRowIndex1 > SLConstants.RowLimit
                                || iColumnIndex1 < 1 || iColumnIndex1 > SLConstants.ColumnLimit)
                            {
                                listFormula[i] = string.Format("{0}{1}", sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty, SLConstants.ErrorReference);
                                HasError = true;
                            }
                            else
                            {
                                listFormula[i] = string.Format("{0}{1}{2}",
                                    sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty,
                                    bIsAbsoluteColumn1 ? "$" + SLTool.ToColumnName(iColumnIndex1) : SLTool.ToColumnName(iColumnIndex1),
                                    bIsAbsoluteRow1 ? "$" + iRowIndex1.ToString(CultureInfo.InvariantCulture) : iRowIndex1.ToString(CultureInfo.InvariantCulture));
                            }
                        }
                    }
                    #endregion end for cell
                }
                else if (iRowDelta != 0 && rgxRow.IsMatch(listFormula[i]))
                {
                    #region for row
                    index = listFormula[i].IndexOf(":");
                    if (index > -1)
                    {
                        sCellRef1 = listFormula[i].Substring(0, index);
                        sCellRef2 = listFormula[i].Substring(index + 1);

                        sSheetName1 = string.Empty;
                        index2 = sCellRef1.IndexOf("!");
                        if (index2 > -1)
                        {
                            sSheetName1 = sCellRef1.Substring(0, index2);
                            sCellRef1 = sCellRef1.Substring(index2 + 1);
                        }

                        sSheetName2 = string.Empty;
                        index2 = sCellRef2.IndexOf("!");
                        if (index2 > -1)
                        {
                            sSheetName2 = sCellRef2.Substring(0, index2);
                            sCellRef2 = sCellRef2.Substring(index + 1);
                        }

                        bIsAbsoluteRow1 = false;
                        bIsAbsoluteRow2 = false;

                        if (sCellRef1.StartsWith("$"))
                        {
                            bIsAbsoluteRow1 = true;
                            sCellRef1 = sCellRef1.Substring(1);
                        }

                        if (sCellRef2.StartsWith("$"))
                        {
                            bIsAbsoluteRow2 = true;
                            sCellRef2 = sCellRef2.Substring(1);
                        }

                        if (int.TryParse(sCellRef1, out iRowIndex1) && int.TryParse(sCellRef2, out iRowIndex2))
                        {
                            // do this once in case the original formula didn't have well-formed range
                            this.SwapRangeIndexIfNecessary(ref sSheetName1, ref sSheetName2, ref iRowIndex1, ref iRowIndex2, ref bIsAbsoluteRow1, ref bIsAbsoluteRow2);

                            bIsCurrentSheet = (string.IsNullOrEmpty(sSheetName1) && string.IsNullOrEmpty(sSheetName2));
                            bIsCurrentSheet |= (gsSelectedWorksheetName.Equals(sSheetName1, StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(sSheetName2));
                            bIsCurrentSheet |= (string.IsNullOrEmpty(sSheetName1) && gsSelectedWorksheetName.Equals(sSheetName2, StringComparison.OrdinalIgnoreCase));
                            bIsCurrentSheet |= (gsSelectedWorksheetName.Equals(sSheetName1, StringComparison.OrdinalIgnoreCase) && gsSelectedWorksheetName.Equals(sSheetName2, StringComparison.OrdinalIgnoreCase));

                            if (ToSwapRowColumn)
                            {
                                // if any absolutes, no swap
                                if (!bIsAbsoluteRow1 && !bIsAbsoluteRow2)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        iRowIndex1 -= StartRowIndex;
                                        iColumnIndex1 = 1 - StartColumnIndex;

                                        iRowIndex2 -= StartRowIndex;
                                        iColumnIndex2 = 1 - StartColumnIndex;

                                        // we set column index as 1 - StartColumnIndex because it's supposed to
                                        // be something like
                                        // iColumnIndex1 -= StartColumnIndex;
                                        // But we don't have a column index here, so we (or Excel rather) take 1
                                        // as the column index.
                                        // In any case, the original code should be
                                        // iColumnIndex1 = -StartColumnIndex
                                        // which is equivalent to
                                        // iColumnIndex1 = 0 - StartColumnIndex

                                        iSwap = iRowIndex1;
                                        iRowIndex1 = iColumnIndex1;
                                        iColumnIndex1 = iSwap;

                                        iSwap = iRowIndex2;
                                        iRowIndex2 = iColumnIndex2;
                                        iColumnIndex2 = iSwap;

                                        iRowIndex1 += AnchorRowIndex;
                                        iColumnIndex1 += AnchorColumnIndex;

                                        iRowIndex2 += AnchorRowIndex;
                                        iColumnIndex2 += AnchorColumnIndex;

                                        if (iRowIndex1 < 1 || iRowIndex1 > SLConstants.RowLimit
                                            || iColumnIndex1 < 1 || iColumnIndex1 > SLConstants.ColumnLimit
                                            || iRowIndex2 < 1 || iRowIndex2 > SLConstants.RowLimit
                                            || iColumnIndex2 < 1 || iColumnIndex2 > SLConstants.ColumnLimit)
                                        {
                                            listFormula[i] = string.Format("{0}{1}", sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty, SLConstants.ErrorReference);
                                            HasError = true;
                                        }
                                        else
                                        {
                                            listFormula[i] = string.Format("{0}{1}${2}:{3}{4}${5}",
                                                sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty,
                                                SLTool.ToColumnName(iColumnIndex1),
                                                iRowIndex1.ToString(CultureInfo.InvariantCulture),
                                                sSheetName2.Length > 0 ? sSheetName2 + "!" : string.Empty,
                                                SLTool.ToColumnName(iColumnIndex2),
                                                SLConstants.RowLimit.ToString(CultureInfo.InvariantCulture));
                                        }
                                    }
                                }
                            }
                            else
                            {
                                bHasCheckError = false;
                                // check for delete range
                                if (CheckForInsertDeleteRowColumn && !CheckForInsert)
                                {
                                    if (CheckForRow)
                                    {
                                        if (CheckStartIndex <= iRowIndex1 && iRowIndex2 <= CheckEndIndex)
                                        {
                                            bHasCheckError = true;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteRow1)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && CheckForRow && iRowDelta > 0)
                                        {
                                            if (iRowIndex1 >= StartRowIndex) iRowIndex1 += iRowDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && CheckForRow && iRowDelta < 0)
                                        {
                                            if (iRowIndex1 >= StartRowIndex) iRowIndex1 += iRowDelta;
                                        }
                                        else
                                        {
                                            iRowIndex1 += iRowDelta;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteRow2)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && CheckForRow && iRowDelta > 0)
                                        {
                                            if (iRowIndex2 >= StartRowIndex) iRowIndex2 += iRowDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && CheckForRow && iRowDelta < 0)
                                        {
                                            if (iRowIndex2 >= StartRowIndex) iRowIndex2 += iRowDelta;
                                        }
                                        else
                                        {
                                            iRowIndex2 += iRowDelta;
                                        }
                                    }
                                }

                                // do this one more time because the delta could've misformed the range
                                this.SwapRangeIndexIfNecessary(ref sSheetName1, ref sSheetName2, ref iRowIndex1, ref iRowIndex2, ref bIsAbsoluteRow1, ref bIsAbsoluteRow2);

                                // override the delta changes above if there's an error.
                                if (bHasCheckError)
                                {
                                    iRowIndex1 = -1;
                                    iRowIndex2 = -1;
                                }

                                if (iRowIndex1 < 1 || iRowIndex1 > SLConstants.RowLimit || iRowIndex2 < 1 || iRowIndex2 > SLConstants.RowLimit)
                                {
                                    listFormula[i] = string.Format("{0}{1}", sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty, SLConstants.ErrorReference);
                                    HasError = true;
                                }
                                else
                                {
                                    listFormula[i] = string.Format("{0}{1}:{2}{3}",
                                        sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty,
                                        bIsAbsoluteRow1 ? "$" + iRowIndex1.ToString(CultureInfo.InvariantCulture) : iRowIndex1.ToString(CultureInfo.InvariantCulture),
                                        sSheetName2.Length > 0 ? sSheetName2 + "!" : string.Empty,
                                        bIsAbsoluteRow2 ? "$" + iRowIndex2.ToString(CultureInfo.InvariantCulture) : iRowIndex2.ToString(CultureInfo.InvariantCulture));
                                }
                            }
                        }
                    }
                    #endregion end for row
                }
                else if (iColumnDelta != 0 && rgxColumn.IsMatch(listFormula[i]))
                {
                    // this is similar to the row version above

                    #region for column
                    index = listFormula[i].IndexOf(":");
                    if (index > -1)
                    {
                        sCellRef1 = listFormula[i].Substring(0, index);
                        sCellRef2 = listFormula[i].Substring(index + 1);

                        sSheetName1 = string.Empty;
                        index2 = sCellRef1.IndexOf("!");
                        if (index2 > -1)
                        {
                            sSheetName1 = sCellRef1.Substring(0, index2);
                            sCellRef1 = sCellRef1.Substring(index2 + 1);
                        }

                        sSheetName2 = string.Empty;
                        index2 = sCellRef2.IndexOf("!");
                        if (index2 > -1)
                        {
                            sSheetName2 = sCellRef2.Substring(0, index2);
                            sCellRef2 = sCellRef2.Substring(index + 1);
                        }

                        bIsAbsoluteColumn1 = false;
                        bIsAbsoluteColumn2 = false;

                        if (sCellRef1.StartsWith("$"))
                        {
                            bIsAbsoluteColumn1 = true;
                            sCellRef1 = sCellRef1.Substring(1);
                        }

                        if (sCellRef2.StartsWith("$"))
                        {
                            bIsAbsoluteColumn2 = true;
                            sCellRef2 = sCellRef2.Substring(1);
                        }

                        iColumnIndex1 = SLTool.ToColumnIndex(sCellRef1);
                        iColumnIndex2 = SLTool.ToColumnIndex(sCellRef2);

                        if (iColumnIndex1 >= 1 && iColumnIndex1 <= SLConstants.ColumnLimit && iColumnIndex2 >= 1 && iColumnIndex2 <= SLConstants.ColumnLimit)
                        {
                            // do this once in case the original formula didn't have well-formed range
                            this.SwapRangeIndexIfNecessary(ref sSheetName1, ref sSheetName2, ref iRowIndex1, ref iRowIndex2, ref bIsAbsoluteRow1, ref bIsAbsoluteRow2);

                            bIsCurrentSheet = (string.IsNullOrEmpty(sSheetName1) && string.IsNullOrEmpty(sSheetName2));
                            bIsCurrentSheet |= (gsSelectedWorksheetName.Equals(sSheetName1, StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(sSheetName2));
                            bIsCurrentSheet |= (string.IsNullOrEmpty(sSheetName1) && gsSelectedWorksheetName.Equals(sSheetName2, StringComparison.OrdinalIgnoreCase));
                            bIsCurrentSheet |= (gsSelectedWorksheetName.Equals(sSheetName1, StringComparison.OrdinalIgnoreCase) && gsSelectedWorksheetName.Equals(sSheetName2, StringComparison.OrdinalIgnoreCase));

                            if (ToSwapRowColumn)
                            {
                                // if any absolutes, no swap
                                if (!bIsAbsoluteColumn1 && !bIsAbsoluteColumn2)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        iRowIndex1 = 1 - StartRowIndex;
                                        iColumnIndex1 -= StartColumnIndex;

                                        iRowIndex2 = 1 - StartRowIndex;
                                        iColumnIndex2 -= StartColumnIndex;

                                        // see comments for row section above

                                        iSwap = iRowIndex1;
                                        iRowIndex1 = iColumnIndex1;
                                        iColumnIndex1 = iSwap;

                                        iSwap = iRowIndex2;
                                        iRowIndex2 = iColumnIndex2;
                                        iColumnIndex2 = iSwap;

                                        iRowIndex1 += AnchorRowIndex;
                                        iColumnIndex1 += AnchorColumnIndex;

                                        iRowIndex2 += AnchorRowIndex;
                                        iColumnIndex2 += AnchorColumnIndex;

                                        if (iRowIndex1 < 1 || iRowIndex1 > SLConstants.RowLimit
                                            || iColumnIndex1 < 1 || iColumnIndex1 > SLConstants.ColumnLimit
                                            || iRowIndex2 < 1 || iRowIndex2 > SLConstants.RowLimit
                                            || iColumnIndex2 < 1 || iColumnIndex2 > SLConstants.ColumnLimit)
                                        {
                                            listFormula[i] = string.Format("{0}{1}", sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty, SLConstants.ErrorReference);
                                            HasError = true;
                                        }
                                        else
                                        {
                                            listFormula[i] = string.Format("{0}${1}{2}:{3}${4}{5}",
                                                sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty,
                                                SLTool.ToColumnName(iColumnIndex1),
                                                iRowIndex1.ToString(CultureInfo.InvariantCulture),
                                                sSheetName2.Length > 0 ? sSheetName2 + "!" : string.Empty,
                                                SLTool.ToColumnName(SLConstants.ColumnLimit),
                                                iRowIndex2.ToString(CultureInfo.InvariantCulture));
                                        }
                                    }
                                }
                            }
                            else
                            {
                                bHasCheckError = false;
                                // check for delete range
                                if (CheckForInsertDeleteRowColumn && !CheckForInsert)
                                {
                                    if (!CheckForRow)
                                    {
                                        if (CheckStartIndex <= iColumnIndex1 && iColumnIndex2 <= CheckEndIndex)
                                        {
                                            bHasCheckError = true;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteColumn1)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && !CheckForRow && iColumnDelta > 0)
                                        {
                                            if (iColumnIndex1 >= StartColumnIndex) iColumnIndex1 += iColumnDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && !CheckForRow && iColumnDelta < 0)
                                        {
                                            if (iColumnIndex1 >= StartColumnIndex) iColumnIndex1 += iColumnDelta;
                                        }
                                        else
                                        {
                                            iColumnIndex1 += iColumnDelta;
                                        }
                                    }
                                }

                                if (!bIsAbsoluteColumn2)
                                {
                                    if ((OnlyForCurrentWorksheet && bIsCurrentSheet) || !OnlyForCurrentWorksheet)
                                    {
                                        if (CheckForInsertDeleteRowColumn && CheckForInsert && !CheckForRow && iColumnDelta > 0)
                                        {
                                            if (iColumnIndex2 >= StartColumnIndex) iColumnIndex2 += iColumnDelta;
                                        }
                                        else if (CheckForInsertDeleteRowColumn && !CheckForInsert && !CheckForRow && iColumnDelta < 0)
                                        {
                                            if (iColumnIndex2 >= StartColumnIndex) iColumnIndex2 += iColumnDelta;
                                        }
                                        else
                                        {
                                            iColumnIndex2 += iColumnDelta;
                                        }
                                    }
                                }

                                // do this one more time because the delta could've misformed the range
                                this.SwapRangeIndexIfNecessary(ref sSheetName1, ref sSheetName2, ref iColumnIndex1, ref iColumnIndex2, ref bIsAbsoluteColumn1, ref bIsAbsoluteColumn2);

                                // override the delta changes above if there's an error.
                                if (bHasCheckError)
                                {
                                    iColumnIndex1 = -1;
                                    iColumnIndex2 = -1;
                                }

                                if (iColumnIndex1 < 1 || iColumnIndex1 > SLConstants.ColumnLimit || iColumnIndex2 < 1 || iColumnIndex2 > SLConstants.ColumnLimit)
                                {
                                    listFormula[i] = string.Format("{0}{1}", sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty, SLConstants.ErrorReference);
                                    HasError = true;
                                }
                                else
                                {
                                    listFormula[i] = string.Format("{0}{1}:{2}{3}",
                                        sSheetName1.Length > 0 ? sSheetName1 + "!" : string.Empty,
                                        bIsAbsoluteColumn1 ? "$" + SLTool.ToColumnName(iColumnIndex1) : SLTool.ToColumnName(iColumnIndex1),
                                        sSheetName2.Length > 0 ? sSheetName2 + "!" : string.Empty,
                                        bIsAbsoluteColumn2 ? "$" + SLTool.ToColumnName(iColumnIndex2) : SLTool.ToColumnName(iColumnIndex2));
                                }
                            }
                        }
                    }
                    #endregion end for column
                }
                // end of ginormous if-elseif-elseif matching the cell, row and column
            }

            StringBuilder sb = new StringBuilder();
            foreach (string s in listFormula)
            {
                sb.Append(s);
            }

            return sb.ToString();
        }

        //... TODO: Delete this whole chunk when the cell formula copying is ok.
        // "Define OK."
        // I don't like you. -_-
        // Maybe wait a couple of minor versions when there are no complains and it seems stable enough...

        /// <summary>
        /// A negative StartRowIndex skips sections of row manipulations.
        /// A negative StartColumnIndex skips sections of column manipulations.
        /// RowDelta and ColumnDelta can be positive or negative
        /// </summary>
        /// <param name="CellFormula"></param>
        /// <param name="StartRowIndex"></param>
        /// <param name="RowDelta"></param>
        /// <param name="StartColumnIndex"></param>
        /// <param name="ColumnDelta"></param>
        /// <returns></returns>
        //internal string AddDeleteCellFormulaDelta(string CellFormula, int StartRowIndex, int RowDelta, int StartColumnIndex, int ColumnDelta)
        //{
        //    string result = string.Empty;
        //    string sToCheck = CellFormula;
        //    string sSheetNameRegex = string.Format("({0}!|'{0}'!)?", gsSelectedWorksheetName);
        //    // This captures A1, A1:B3, Sheet1!A1, 'Sheet1'!A1, B2:Sheet1!C4, Sheet1!B2:C4 and so on.
        //    // Basically it captures single cell references (A1) and cell ranges (A1:B3).
        //    // It also captures the worksheet name too.
        //    // We use the selected worksheet name in the regex because we're only interested in
        //    // modifying any cell references/ranges on the selected worksheet.
        //    // This automatically limit the regex matches to those we want.
        //    // Note that we only care for the ":" as the range character. Apparently, Excel
        //    // accepts A1.B2 as a valid range, but auto-corrects it to A1:B2 immediately.
        //    // Otherwise, we could use \s*[:.]\s* and then we have to handle the case with
        //    // the period . as the range character in the post-processing.
        //    string sCellRefRegex = @"(?<cellref>" + sSheetNameRegex + @"\$?[a-zA-Z]{1,3}\$?[0-9]{1,7}(\s*:\s*" + sSheetNameRegex + @"\$?[a-zA-Z]{1,3}\$?[0-9]{1,7})?)";
        //    // The only characters that can be before a valid cell reference are +-*/^=<>,( and the space.
        //    // The cell reference can also be at the start of the string, thus the ^
        //    // The only characters that can be after a valid cell reference are +-*/^=<>,) and the space.
        //    // The cell reference can also be at the end of the string, thus the $
        //    string sRegexCheck = @"(?<cellrefpre>^|[+\-*/^=<>,(]|\s)" + sCellRefRegex + @"(?<cellrefpost>[+\-*/^=<>,)]|\s|$)";
        //    int index = 0;
        //    int iDoubleQuoteCount = 0;
        //    Match m;
        //    m = Regex.Match(sToCheck, sRegexCheck);
        //    while (m.Success)
        //    {
        //        index = sToCheck.IndexOf(m.Value);
        //        result += sToCheck.Substring(0, index) + m.Groups["cellrefpre"].Value;
        //        sToCheck = sToCheck.Substring(index + m.Value.Length);

        //        iDoubleQuoteCount = result.Length - result.Replace("\"", "").Length;
        //        // This checks if there's a matching pair of double quotes.
        //        // If there's an odd number of double quotes, then the matched
        //        // value is behind a double quote, and hence should be taken
        //        // as a literal string.
        //        if (iDoubleQuoteCount % 2 == 0)
        //        {
        //            result += AddDeleteCellReferenceDelta(m.Groups["cellref"].Value, StartRowIndex, RowDelta, StartColumnIndex, ColumnDelta);
        //        }
        //        else
        //        {
        //            result += m.Groups["cellref"].Value;
        //        }
        //        result += m.Groups["cellrefpost"].Value;

        //        m = Regex.Match(sToCheck, sRegexCheck);
        //    }
        //    result += sToCheck;

        //    return result;
        //}

        /// <summary>
        /// This closely follows the logic of AddDeleteCellFormulaDelta()
        /// Delta can be positive or negative.
        /// </summary>
        /// <param name="DefinedNameValue"></param>
        /// <param name="CheckForRow"></param>
        /// <param name="StartRange"></param>
        /// <param name="Delta"></param>
        /// <returns></returns>
        //internal string AddDeleteDefinedNameRowColumnRangeDelta(string DefinedNameValue, bool CheckForRow, int StartRange, int Delta)
        //{
        //    string result = string.Empty;
        //    string sToCheck = DefinedNameValue;
        //    string sSheetNameRegex = string.Format("({0}!|'{0}'!)?", gsSelectedWorksheetName);
        //    // We want to capture strings such as Sheet1!$B:$D or Sheet1!$3:$9
        //    // In this case, we only care about the $ for the "absolute-ness"
        //    // While Sheet1!3:5 may be a valid defined name value (I don't know...),
        //    // we will ignore that because it's a relative reference.
        //    string sCellRefRegex;
        //    if (CheckForRow)
        //    {
        //        sCellRefRegex = @"(?<cellref>" + sSheetNameRegex + @"\$[0-9]{1,7}\s*:\s*" + sSheetNameRegex + @"\$[0-9]{1,7})";
        //    }
        //    else
        //    {
        //        sCellRefRegex = @"(?<cellref>" + sSheetNameRegex + @"\$[a-zA-Z]{1,3}\s*:\s*" + sSheetNameRegex + @"\$[a-zA-Z]{1,3})";
        //    }
        //    // The only characters that can be before a valid cell reference are +-*/^=<>,( and the space.
        //    // The cell reference can also be at the start of the string, thus the ^
        //    // The only characters that can be after a valid cell reference are +-*/^=<>,) and the space.
        //    // The cell reference can also be at the end of the string, thus the $
        //    string sRegexCheck = @"(?<cellrefpre>^|[+\-*/^=<>,(]|\s)" + sCellRefRegex + @"(?<cellrefpost>[+\-*/^=<>,)]|\s|$)";
        //    int index = 0;
        //    int iDoubleQuoteCount = 0;
        //    Match m;
        //    m = Regex.Match(sToCheck, sRegexCheck);
        //    while (m.Success)
        //    {
        //        index = sToCheck.IndexOf(m.Value);
        //        result += sToCheck.Substring(0, index) + m.Groups["cellrefpre"].Value;
        //        sToCheck = sToCheck.Substring(index + m.Value.Length);

        //        iDoubleQuoteCount = result.Length - result.Replace("\"", "").Length;
        //        // This checks if there's a matching pair of double quotes.
        //        // If there's an odd number of double quotes, then the matched
        //        // value is behind a double quote, and hence should be taken
        //        // as a literal string.
        //        if (iDoubleQuoteCount % 2 == 0)
        //        {
        //            result += AddDeleteRowColumnRangeDelta(m.Groups["cellref"].Value, CheckForRow, StartRange, Delta);
        //        }
        //        else
        //        {
        //            result += m.Groups["cellref"].Value;
        //        }
        //        result += m.Groups["cellrefpost"].Value;

        //        m = Regex.Match(sToCheck, sRegexCheck);
        //    }
        //    result += sToCheck;

        //    return result;
        //}

        /// <summary>
        /// Delta can be positive or negative
        /// </summary>
        /// <param name="Range"></param>
        /// <param name="CheckForRow"></param>
        /// <param name="StartRange"></param>
        /// <param name="Delta"></param>
        /// <returns></returns>
        //internal string AddDeleteRowColumnRangeDelta(string Range, bool CheckForRow, int StartRange, int Delta)
        //{
        //    string result = string.Empty;
        //    string sSheetName = string.Empty, sSheetName2 = string.Empty;
        //    string sRef1 = string.Empty, sRef2 = string.Empty;
        //    int iRowIndex = -1, iColumnIndex = -1;
        //    int iRowIndex2 = -1, iColumnIndex2 = -1;
        //    int iEndRange = -1;
        //    int index = 0;
        //    index = Range.LastIndexOf(":");
        //    if (index < 0)
        //    {
        //        // this case shouldn't happen...
        //        result = Range;
        //    }
        //    else
        //    {
        //        sSheetName = Range.Substring(0, index).Trim();
        //        sSheetName2 = Range.Substring(index + 1).Trim();

        //        index = sSheetName.LastIndexOf("!");
        //        if (index < 0)
        //        {
        //            sRef1 = sSheetName.Replace("$", "").Trim();
        //            sSheetName = string.Empty;
        //        }
        //        else
        //        {
        //            sRef1 = sSheetName.Substring(index + 1).Replace("$", "").Trim();
        //            sSheetName = sSheetName.Substring(0, index + 1);
        //        }

        //        index = sSheetName2.LastIndexOf("!");
        //        if (index < 0)
        //        {
        //            sRef2 = sSheetName2.Replace("$", "").Trim();
        //            sSheetName2 = string.Empty;
        //        }
        //        else
        //        {
        //            sRef2 = sSheetName2.Substring(index + 1).Replace("$", "").Trim();
        //            sSheetName2 = sSheetName2.Substring(0, index + 1);
        //        }

        //        if (Delta >= 0)
        //        {
        //            iEndRange = StartRange + Delta;
        //        }
        //        else
        //        {
        //            iEndRange = StartRange - Delta - 1;
        //        }

        //        if (CheckForRow)
        //        {
        //            if (int.TryParse(sRef1, out iRowIndex) && int.TryParse(sRef2, out iRowIndex2))
        //            {
        //                if (Delta >= 0)
        //                {
        //                    AddRowColumnIndexDelta(StartRange, Delta, true, ref iRowIndex, ref iRowIndex2);
        //                }
        //                else
        //                {
        //                    if (StartRange <= iRowIndex && iRowIndex2 <= iEndRange)
        //                    {
        //                        iRowIndex = -1;
        //                        iRowIndex2 = -1;
        //                    }
        //                    else
        //                    {
        //                        DeleteRowColumnIndexDelta(StartRange, iEndRange, -Delta, ref iRowIndex, ref iRowIndex2);
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                iRowIndex = -1;
        //                iRowIndex2 = -1;
        //            }

        //            if (iRowIndex < 1 || iRowIndex > SLConstants.RowLimit || iRowIndex2 < 1 || iRowIndex2 > SLConstants.RowLimit)
        //            {
        //                result = sSheetName + "#REF!";
        //            }
        //            else
        //            {
        //                result = string.Format("{0}${1}:{2}${3}", sSheetName, iRowIndex.ToString(CultureInfo.InvariantCulture), sSheetName2, iRowIndex2.ToString(CultureInfo.InvariantCulture));
        //            }
        //        }
        //        else
        //        {
        //            iColumnIndex = SLTool.ToColumnIndex(sRef1);
        //            iColumnIndex2 = SLTool.ToColumnIndex(sRef2);
        //            if (iColumnIndex > 0 && iColumnIndex2 > 0)
        //            {
        //                if (Delta >= 0)
        //                {
        //                    AddRowColumnIndexDelta(StartRange, Delta, false, ref iColumnIndex, ref iColumnIndex2);
        //                }
        //                else
        //                {
        //                    if (StartRange <= iColumnIndex && iColumnIndex2 <= iEndRange)
        //                    {
        //                        iColumnIndex = -1;
        //                        iColumnIndex2 = -1;
        //                    }
        //                    else
        //                    {
        //                        DeleteRowColumnIndexDelta(StartRange, iEndRange, -Delta, ref iColumnIndex, ref iColumnIndex2);
        //                    }
        //                }
        //            }

        //            if (iColumnIndex < 1 || iColumnIndex > SLConstants.ColumnLimit || iColumnIndex2 < 1 || iColumnIndex2 > SLConstants.ColumnLimit)
        //            {
        //                result = sSheetName + "#REF!";
        //            }
        //            else
        //            {
        //                result = string.Format("{0}${1}:{2}${3}", sSheetName, SLTool.ToColumnName(iColumnIndex), sSheetName2, SLTool.ToColumnName(iColumnIndex2));
        //            }
        //        }
        //    }

        //    return result;
        //}

        /// <summary>
        /// A negative StartRowIndex skips sections of row manipulations.
        /// A negative StartColumnIndex skips sections of column manipulations.
        /// RowDelta and ColumnDelta can be positive or negative.
        /// </summary>
        /// <param name="CellReference"></param>
        /// <param name="StartRowIndex"></param>
        /// <param name="RowDelta"></param>
        /// <param name="StartColumnIndex"></param>
        /// <param name="ColumnDelta"></param>
        /// <returns></returns>
        //internal string AddDeleteCellReferenceDelta(string CellReference, int StartRowIndex, int RowDelta, int StartColumnIndex, int ColumnDelta)
        //{
        //    string result = string.Empty;
        //    string sSheetName = string.Empty, sSheetName2 = string.Empty;
        //    string sCellRef = string.Empty, sCellRef2 = string.Empty;
        //    bool bIsRange = false;
        //    int index = 0;
        //    index = CellReference.LastIndexOf(":");
        //    if (index < 0)
        //    {
        //        bIsRange = false;
        //        index = CellReference.LastIndexOf("!");
        //        if (index < 0)
        //        {
        //            sSheetName = string.Empty;
        //            sCellRef = CellReference.Trim();
        //        }
        //        else
        //        {
        //            sSheetName = CellReference.Substring(0, index).Trim() + "!";
        //            sCellRef = CellReference.Substring(index + 1).Trim();
        //        }
        //        sSheetName2 = string.Empty;
        //        sCellRef2 = string.Empty;
        //    }
        //    else
        //    {
        //        bIsRange = true;
        //        sCellRef = CellReference.Substring(0, index);
        //        sCellRef2 = CellReference.Substring(index + 1);

        //        index = sCellRef.LastIndexOf("!");
        //        if (index < 0)
        //        {
        //            sSheetName = string.Empty;
        //            sCellRef = sCellRef.Trim();
        //        }
        //        else
        //        {
        //            sSheetName = sCellRef.Substring(0, index).Trim() + "!";
        //            sCellRef = sCellRef.Substring(index + 1).Trim();
        //        }

        //        index = sCellRef2.LastIndexOf("!");
        //        if (index < 0)
        //        {
        //            sSheetName2 = string.Empty;
        //            sCellRef2 = sCellRef2.Trim();
        //        }
        //        else
        //        {
        //            sSheetName2 = sCellRef2.Substring(0, index).Trim() + "!";
        //            sCellRef2 = sCellRef2.Substring(index + 1).Trim();
        //        }
        //    }

        //    bool bIsRowAbsolute = Regex.IsMatch(sCellRef, @"\$[0-9]{1,7}");
        //    bool bIsColumnAbsolute = Regex.IsMatch(sCellRef, @"\$[a-zA-Z]{1,3}");
        //    bool bIsRowAbsolute2 = false, bIsColumnAbsolute2 = false;
        //    sCellRef = sCellRef.Replace("$", "");
        //    if (bIsRange)
        //    {
        //        bIsRowAbsolute2 = Regex.IsMatch(sCellRef2, @"\$[0-9]{1,7}");
        //        bIsColumnAbsolute2 = Regex.IsMatch(sCellRef2, @"\$[a-zA-Z]{1,3}");
        //        sCellRef2 = sCellRef2.Replace("$", "");
        //    }
        //    int iRowIndex = -1, iColumnIndex = -1;
        //    int iRowIndex2 = -1, iColumnIndex2 = -1;
        //    int iEndRowIndex = -1, iEndColumnIndex = -1;

        //    if (RowDelta >= 0)
        //    {
        //        iEndRowIndex = StartRowIndex + RowDelta;
        //    }
        //    else
        //    {
        //        iEndRowIndex = StartRowIndex - RowDelta - 1;
        //    }

        //    if (ColumnDelta >= 0)
        //    {
        //        iEndColumnIndex = StartColumnIndex + ColumnDelta;
        //    }
        //    else
        //    {
        //        iEndColumnIndex = StartColumnIndex - ColumnDelta - 1;
        //    }

        //    result = CellReference;
        //    if (!bIsRange)
        //    {
        //        if (SLTool.FormatCellReferenceToRowColumnIndex(sCellRef, out iRowIndex, out iColumnIndex))
        //        {
        //            if (StartRowIndex > 0)
        //            {
        //                if (RowDelta > 0)
        //                {
        //                    if (iRowIndex >= StartRowIndex)
        //                    {
        //                        iRowIndex += RowDelta;
        //                    }
        //                }
        //                else
        //                {
        //                    if (StartRowIndex <= iRowIndex && iRowIndex <= iEndRowIndex)
        //                    {
        //                        iRowIndex = -1;
        //                    }
        //                    else if (iEndRowIndex < iRowIndex)
        //                    {
        //                        // the delta is negative, so add it
        //                        iRowIndex += RowDelta;
        //                    }
        //                }
        //            }

        //            if (StartColumnIndex > 0)
        //            {
        //                if (ColumnDelta > 0)
        //                {
        //                    if (iColumnIndex >= StartColumnIndex)
        //                    {
        //                        iColumnIndex += ColumnDelta;
        //                    }
        //                }
        //                else
        //                {
        //                    if (StartColumnIndex <= iColumnIndex && iColumnIndex <= iEndColumnIndex)
        //                    {
        //                        iColumnIndex = -1;
        //                    }
        //                    else if (iEndColumnIndex < iColumnIndex)
        //                    {
        //                        // the delta is negative, so add it
        //                        iColumnIndex += ColumnDelta;
        //                    }
        //                }
        //            }

        //            if (iRowIndex < 1 || iRowIndex > SLConstants.RowLimit || iColumnIndex < 1 || iColumnIndex > SLConstants.ColumnLimit)
        //            {
        //                result = "#REF!";
        //            }
        //            else
        //            {
        //                // would the cell references be independently absolute or relative?
        //                // Otherwise we'd use SLTool to form the cell reference...
        //                result = sSheetName + (bIsColumnAbsolute ? "$" : "") + SLTool.ToColumnName(iColumnIndex) + (bIsRowAbsolute ? "$" : "") + iRowIndex.ToString(CultureInfo.InvariantCulture);
        //            }
        //        }
        //    }
        //    else
        //    {
        //        if (SLTool.FormatCellReferenceToRowColumnIndex(sCellRef, out iRowIndex, out iColumnIndex) && SLTool.FormatCellReferenceToRowColumnIndex(sCellRef2, out iRowIndex2, out iColumnIndex2))
        //        {
        //            if (StartRowIndex > 0)
        //            {
        //                if (RowDelta > 0)
        //                {
        //                    AddRowColumnIndexDelta(StartRowIndex, RowDelta, true, ref iRowIndex, ref iRowIndex2);
        //                }
        //                else
        //                {
        //                    if (StartRowIndex <= iRowIndex && iRowIndex2 <= iEndRowIndex)
        //                    {
        //                        iRowIndex = -1;
        //                        iRowIndex2 = -1;
        //                    }
        //                    else
        //                    {
        //                        DeleteRowColumnIndexDelta(StartRowIndex, iEndRowIndex, -RowDelta, ref iRowIndex, ref iRowIndex2);
        //                    }
        //                }
        //            }

        //            if (StartColumnIndex > 0)
        //            {
        //                if (ColumnDelta > 0)
        //                {
        //                    AddRowColumnIndexDelta(StartColumnIndex, ColumnDelta, false, ref iColumnIndex, ref iColumnIndex2);
        //                }
        //                else
        //                {
        //                    if (StartColumnIndex <= iColumnIndex && iColumnIndex2 <= iEndColumnIndex)
        //                    {
        //                        iColumnIndex = -1;
        //                        iColumnIndex2 = -1;
        //                    }
        //                    else
        //                    {
        //                        DeleteRowColumnIndexDelta(StartColumnIndex, iEndColumnIndex, -ColumnDelta, ref iColumnIndex, ref iColumnIndex2);
        //                    }
        //                }
        //            }

        //            if (iRowIndex < 1 || iRowIndex > SLConstants.RowLimit || iColumnIndex < 1 || iColumnIndex > SLConstants.ColumnLimit || iRowIndex2 < 1 || iRowIndex2 > SLConstants.RowLimit || iColumnIndex2 < 1 || iColumnIndex2 > SLConstants.ColumnLimit)
        //            {
        //                result = "#REF!";
        //            }
        //            else
        //            {
        //                // would the cell references be independently absolute or relative?
        //                // Otherwise we'd use SLTool to form the cell reference...
        //                result = sSheetName + (bIsColumnAbsolute ? "$" : "") + SLTool.ToColumnName(iColumnIndex) + (bIsRowAbsolute ? "$" : "") + iRowIndex.ToString(CultureInfo.InvariantCulture);
        //                result += ":" + sSheetName2 + (bIsColumnAbsolute2 ? "$" : "") + SLTool.ToColumnName(iColumnIndex2) + (bIsRowAbsolute2 ? "$" : "") + iRowIndex2.ToString(CultureInfo.InvariantCulture);
        //            }
        //        }
        //    }

        //    return result;
        //}
    }
}
