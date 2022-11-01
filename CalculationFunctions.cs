using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// Flattens shared cell formulas into respective cells.
        /// For example, if there's a shared cell formula in A2 for the range A2:A6, the shared cell formula will be
        /// individually assigned into A2, A3, A4, A5 and A6. And the shared cell formula portion will then be removed.
        /// </summary>
        public void FlattenAllSharedCellFormula()
        {
            if (slws.SharedCellFormulas.Count > 0)
            {
                SLCell cell;
                int i, iRowIndex, iColumnIndex;
                bool bHasError;
                foreach (SLSharedCellFormula scf in slws.SharedCellFormulas.Values)
                {
                    for (i = 0; i < scf.Reference.Count; ++i)
                    {
                        for (iRowIndex = scf.Reference[i].StartRowIndex; iRowIndex <= scf.Reference[i].EndRowIndex; ++iRowIndex)
                        {
                            for (iColumnIndex = scf.Reference[i].StartColumnIndex; iColumnIndex <= scf.Reference[i].EndColumnIndex; ++iColumnIndex)
                            {
                                if (slws.CellWarehouse.Exists(iRowIndex, iColumnIndex))
                                {
                                    cell = slws.CellWarehouse.Cells[iRowIndex][iColumnIndex].Clone();
                                    if (iRowIndex == scf.BaseCellRowIndex && iColumnIndex == scf.BaseCellColumnIndex)
                                    {
                                        cell.CellFormula = new SLCellFormula();
                                        cell.CellFormula.FormulaType = CellFormulaValues.Normal;
                                        cell.CellFormula.FormulaText = scf.FormulaText;
                                        cell.CellText = "";
                                    }
                                    else
                                    {
                                        bHasError = false;
                                        cell.CellFormula = new SLCellFormula();
                                        cell.CellFormula.FormulaType = CellFormulaValues.Normal;
                                        cell.CellFormula.FormulaText = AdjustCellFormulaDelta(scf.FormulaText, false, scf.BaseCellRowIndex, scf.BaseCellColumnIndex, iRowIndex, iColumnIndex, false, false, false, false, 0, 0, out bHasError);
                                        if (bHasError)
                                        {
                                            cell.CellText = SLConstants.ErrorReference;
                                            cell.DataType = CellValues.Error;
                                        }
                                        else
                                        {
                                            cell.CellText = "";
                                        }
                                    }

                                    slws.CellWarehouse.SetValue(iRowIndex, iColumnIndex, cell);
                                }
                            }
                        }
                    }
                }

                slws.SharedCellFormulas.Clear();
            }
        }

        internal bool Calculate(TotalsRowFunctionValues Function, List<SLCell> Cells, out string ResultText)
        {
            if (Function == TotalsRowFunctionValues.None)
            {
                ResultText = string.Empty;
                return true;
            }

            SLDataFieldFunctionValues func = SLDataFieldFunctionValues.Sum;
            switch (Function)
            {
                case TotalsRowFunctionValues.Average:
                    func = SLDataFieldFunctionValues.Average;
                    break;
                case TotalsRowFunctionValues.Count:
                    func = SLDataFieldFunctionValues.Count;
                    break;
                case TotalsRowFunctionValues.CountNumbers:
                    func = SLDataFieldFunctionValues.CountNumbers;
                    break;
                case TotalsRowFunctionValues.Maximum:
                    func = SLDataFieldFunctionValues.Maximum;
                    break;
                case TotalsRowFunctionValues.Minimum:
                    func = SLDataFieldFunctionValues.Minimum;
                    break;
                case TotalsRowFunctionValues.StandardDeviation:
                    func = SLDataFieldFunctionValues.StandardDeviation;
                    break;
                case TotalsRowFunctionValues.Sum:
                    func = SLDataFieldFunctionValues.Sum;
                    break;
                case TotalsRowFunctionValues.Variance:
                    func = SLDataFieldFunctionValues.Variance;
                    break;
            }

            return Calculate(func, Cells, out ResultText);
        }

        internal bool Calculate(SLDataFieldFunctionValues Function, List<SLCell> Cells, out string ResultText)
        {
            bool result = false;
            ResultText = string.Empty;

            int i;
            int iCount = 0;
            double fTemp = 0;
            double fValue = 0;
            double fMean = 0;
            List<double> listMean = new List<double>();
            bool bFound = false;

            switch (Function)
            {
                case SLDataFieldFunctionValues.Average:
                    iCount = 0;
                    fTemp = 0.0;
                    foreach (SLCell c in Cells)
                    {
                        if (c.DataType == CellValues.Number)
                        {
                            if (c.CellText != null)
                            {
                                if (double.TryParse(c.CellText, out fValue))
                                {
                                    ++iCount;
                                    fTemp += fValue;
                                }
                            }
                            else
                            {
                                fValue = c.NumericValue;
                                ++iCount;
                                fTemp += fValue;
                            }
                        }
                    }

                    if (iCount == 0)
                    {
                        result = false;
                        ResultText = SLConstants.ErrorDivisionByZero;
                    }
                    else
                    {
                        result = true;
                        fTemp = fTemp / iCount;
                        ResultText = fTemp.ToString(CultureInfo.InvariantCulture);
                    }
                    break;
                case SLDataFieldFunctionValues.Count:
                    iCount = 0;
                    foreach (SLCell c in Cells)
                    {
                        if (c.CellText != null)
                        {
                            ++iCount;
                        }
                        else
                        {
                            if (c.DataType == CellValues.Number || c.DataType == CellValues.SharedString || c.DataType == CellValues.Boolean)
                            {
                                ++iCount;
                            }
                        }
                    }

                    result = true;
                    ResultText = iCount.ToString(CultureInfo.InvariantCulture);
                    break;
                case SLDataFieldFunctionValues.CountNumbers:
                    iCount = 0;
                    foreach (SLCell c in Cells)
                    {
                        // we're not going to check the cell value itself...
                        if (c.DataType == CellValues.Number) ++iCount;
                    }

                    result = true;
                    ResultText = iCount.ToString(CultureInfo.InvariantCulture);
                    break;
                case SLDataFieldFunctionValues.Maximum:
                    bFound = false;
                    fTemp = double.NegativeInfinity;
                    foreach (SLCell c in Cells)
                    {
                        if (c.DataType == CellValues.Number)
                        {
                            if (c.CellText != null)
                            {
                                if (double.TryParse(c.CellText, out fValue))
                                {
                                    bFound = true;
                                    if (fValue > fTemp) fTemp = fValue;
                                }
                            }
                            else
                            {
                                bFound = true;
                                if (c.NumericValue > fTemp) fTemp = c.NumericValue;
                            }
                        }
                    }

                    result = true;
                    ResultText = bFound ? fTemp.ToString(CultureInfo.InvariantCulture) : "0";
                    break;
                case SLDataFieldFunctionValues.Minimum:
                    bFound = false;
                    fTemp = double.PositiveInfinity;
                    foreach (SLCell c in Cells)
                    {
                        if (c.DataType == CellValues.Number)
                        {
                            if (c.CellText != null)
                            {
                                if (double.TryParse(c.CellText, out fValue))
                                {
                                    bFound = true;
                                    if (fValue < fTemp) fTemp = fValue;
                                }
                            }
                            else
                            {
                                bFound = true;
                                if (c.NumericValue < fTemp) fTemp = c.NumericValue;
                            }
                        }
                    }

                    result = true;
                    ResultText = bFound ? fTemp.ToString(CultureInfo.InvariantCulture) : "0";
                    break;
                case SLDataFieldFunctionValues.Product:
                    fTemp = 1.0;
                    foreach (SLCell c in Cells)
                    {
                        if (c.DataType == CellValues.Number)
                        {
                            if (c.CellText != null)
                            {
                                if (double.TryParse(c.CellText, out fValue))
                                {
                                    fTemp *= fValue;
                                }
                            }
                            else
                            {
                                fTemp *= c.NumericValue;
                            }
                        }
                    }

                    result = true;
                    ResultText = fTemp.ToString(CultureInfo.InvariantCulture);
                    break;
                case SLDataFieldFunctionValues.StandardDeviation:
                    iCount = 0;
                    fTemp = 0.0;
                    listMean = new List<double>();
                    foreach (SLCell c in Cells)
                    {
                        if (c.DataType == CellValues.Number)
                        {
                            if (c.CellText != null)
                            {
                                if (double.TryParse(c.CellText, out fValue))
                                {
                                    ++iCount;
                                    fTemp += fValue;
                                    listMean.Add(fValue);
                                }
                            }
                            else
                            {
                                ++iCount;
                                fTemp += c.NumericValue;
                                listMean.Add(c.NumericValue);
                            }
                        }
                    }

                    if (iCount > 0)
                    {
                        fMean = fTemp / iCount;
                        fTemp = 0.0;
                        for (i = 0; i < listMean.Count; ++i)
                        {
                            fTemp += ((fMean - listMean[i]) * (fMean - listMean[i]));
                        }
                        fTemp = Math.Sqrt(fTemp / iCount);

                        result = true;
                        ResultText = fTemp.ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        result = false;
                        ResultText = SLConstants.ErrorDivisionByZero;
                    }
                    break;
                case SLDataFieldFunctionValues.Sum:
                    fTemp = 0.0;
                    foreach (SLCell c in Cells)
                    {
                        if (c.DataType == CellValues.Number)
                        {
                            if (c.CellText != null)
                            {
                                if (double.TryParse(c.CellText, out fValue))
                                {
                                    fTemp += fValue;
                                }
                            }
                            else
                            {
                                fTemp += c.NumericValue;
                            }
                        }
                    }

                    result = true;
                    ResultText = fTemp.ToString(CultureInfo.InvariantCulture);
                    break;
                case SLDataFieldFunctionValues.Variance:
                    iCount = 0;
                    fTemp = 0.0;
                    fMean = 0.0;
                    listMean = new List<double>();
                    foreach (SLCell c in Cells)
                    {
                        if (c.DataType == CellValues.Number)
                        {
                            if (c.CellText != null)
                            {
                                if (double.TryParse(c.CellText, out fValue))
                                {
                                    ++iCount;
                                    fMean += fValue;
                                    fTemp += (fValue * fValue);
                                }
                            }
                            else
                            {
                                ++iCount;
                                fMean += c.NumericValue;
                                fTemp += (c.NumericValue * c.NumericValue);
                            }
                        }
                    }

                    if (iCount <= 1)
                    {
                        result = false;
                        ResultText = SLConstants.ErrorDivisionByZero;
                    }
                    else
                    {
                        result = true;
                        --iCount;
                        fTemp = (fMean / iCount) - ((fTemp / iCount) * (fTemp / iCount));
                        ResultText = fTemp.ToString(CultureInfo.InvariantCulture);
                    }
                    break;
            }

            return result;
        }

        internal int GetFunctionNumber(TotalsRowFunctionValues Function)
        {
            int result = 0;
            switch (Function)
            {
                case TotalsRowFunctionValues.Average:
                    result = 101;
                    break;
                case TotalsRowFunctionValues.Count:
                    result = 103;
                    break;
                case TotalsRowFunctionValues.CountNumbers:
                    result = 102;
                    break;
                case TotalsRowFunctionValues.Maximum:
                    result = 104;
                    break;
                case TotalsRowFunctionValues.Minimum:
                    result = 105;
                    break;
                case TotalsRowFunctionValues.StandardDeviation:
                    result = 107;
                    break;
                case TotalsRowFunctionValues.Sum:
                    result = 109;
                    break;
                case TotalsRowFunctionValues.Variance:
                    result = 110;
                    break;
            }

            return result;
        }
    }
}
