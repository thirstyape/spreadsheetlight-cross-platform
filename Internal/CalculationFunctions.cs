using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace SpreadsheetLight;

internal static class CalculationFunctions
{
	internal static void FlattenAllSharedCellFormula(SLDocument document)
	{
		if (document.slws.SharedCellFormulas.Count == 0)
			return;

		foreach (var formula in document.slws.SharedCellFormulas.Values)
		{
			for (var i = 0; i < formula.Reference.Count; ++i)
			{
				for (var row = formula.Reference[i].StartRowIndex; row <= formula.Reference[i].EndRowIndex; ++row)
				{
					for (var column = formula.Reference[i].StartColumnIndex; column <= formula.Reference[i].EndColumnIndex; ++column)
					{
						if (document.slws.CellWarehouse.Exists(row, column) == false)
							continue;

						var cell = document.slws.CellWarehouse.Cells[row][column].Clone();

						if (row == formula.BaseCellRowIndex && column == formula.BaseCellColumnIndex)
						{
							cell.CellFormula = new()
							{
								FormulaType = CellFormulaValues.Normal,
								FormulaText = formula.FormulaText
							};

							cell.CellText = "";
						}
						else
						{
							cell.CellFormula = new()
							{
								FormulaType = CellFormulaValues.Normal,
								FormulaText = document.AdjustCellFormulaDelta(formula.FormulaText, false, formula.BaseCellRowIndex, formula.BaseCellColumnIndex, row, column, false, false, false, false, 0, 0, out bool error)
							};

							if (error)
							{
								cell.CellText = SLConstants.ErrorReference;
								cell.DataType = CellValues.Error;
							}
							else
							{
								cell.CellText = "";
							}
						}

						document.slws.CellWarehouse.SetValue(row, column, cell);
					}
				}
			}
		}

		document.slws.SharedCellFormulas.Clear();
	}

	internal static bool Calculate(TotalsRowFunctionValues function, List<SLCell> cells, out string resultText)
	{
		if (function == TotalsRowFunctionValues.None)
		{
			resultText = string.Empty;
			return true;
		}

		return function switch
		{
			TotalsRowFunctionValues.Minimum => Calculate(SLDataFieldFunctionValues.Minimum, cells, out resultText),
			TotalsRowFunctionValues.Maximum => Calculate(SLDataFieldFunctionValues.Maximum, cells, out resultText),
			TotalsRowFunctionValues.Average => Calculate(SLDataFieldFunctionValues.Average, cells, out resultText),
			TotalsRowFunctionValues.Count => Calculate(SLDataFieldFunctionValues.Count, cells, out resultText),
			TotalsRowFunctionValues.CountNumbers => Calculate(SLDataFieldFunctionValues.CountNumbers, cells, out resultText),
			TotalsRowFunctionValues.StandardDeviation => Calculate(SLDataFieldFunctionValues.StandardDeviation, cells, out resultText),
			TotalsRowFunctionValues.Variance => Calculate(SLDataFieldFunctionValues.Variance, cells, out resultText),
			_ => Calculate(SLDataFieldFunctionValues.Sum, cells, out resultText)
		};
	}

	internal static bool Calculate(SLDataFieldFunctionValues function, List<SLCell> cells, out string resultText)
	{
		double temp, value, mean;

		var matched = false;
		var count = 0;
		var success = false;
		var means = new List<double>();

		resultText = string.Empty;

		switch (function)
		{
			case SLDataFieldFunctionValues.Average:
				temp = 0D;

				foreach (var c in cells.Where(c => c.DataType == CellValues.Number))
				{
					if (c.CellText == null)
					{
						value = c.NumericValue;
						++count;
						temp += value;
					}
					else if (double.TryParse(c.CellText, out value))
					{
						++count;
						temp += value;
					}
				}

				success = count != 0;

				if (count == 0)
				{
					resultText = SLConstants.ErrorDivisionByZero;
				}
				else
				{
					temp /= count;
					resultText = temp.ToString(CultureInfo.InvariantCulture);
				}

				break;
			case SLDataFieldFunctionValues.Count:
				success = true;

				resultText = cells
					.Where(c => c.CellText != null || c.DataType == CellValues.Number || c.DataType == CellValues.SharedString || c.DataType == CellValues.Boolean)
					.Count()
					.ToString(CultureInfo.InvariantCulture);

				resultText = count.ToString(CultureInfo.InvariantCulture);

				break;
			case SLDataFieldFunctionValues.CountNumbers:
				success = true;

				resultText = cells
					.Where(c => c.DataType == CellValues.Number)
					.Count()
					.ToString(CultureInfo.InvariantCulture);

				break;
			case SLDataFieldFunctionValues.Maximum:
				temp = double.NegativeInfinity;

				foreach (var c in cells.Where(c => c.DataType == CellValues.Number))
				{
					if (c.CellText == null)
					{
						matched = true;

						if (c.NumericValue > temp)
							temp = c.NumericValue;
					}
					else if (double.TryParse(c.CellText, out value))
					{
						matched = true;

						if (value > temp)
							temp = value;
					}
				}

				success = true;
				resultText = matched ? temp.ToString(CultureInfo.InvariantCulture) : "0";
				break;
			case SLDataFieldFunctionValues.Minimum:
				temp = double.PositiveInfinity;

				foreach (var c in cells.Where(c => c.DataType == CellValues.Number))
				{
					if (c.CellText == null)
					{
						matched = true;

						if (c.NumericValue < temp)
							temp = c.NumericValue;
					}
					else if (double.TryParse(c.CellText, out value))
					{
						matched = true;

						if (value < temp)
							temp = value;
					}
				}

				success = true;
				resultText = matched ? temp.ToString(CultureInfo.InvariantCulture) : "0";
				break;
			case SLDataFieldFunctionValues.Product:
				temp = 1D;

				foreach (var c in cells.Where(c => c.DataType == CellValues.Number))
				{
					if (c.CellText == null)
						temp *= c.NumericValue;
					else if (double.TryParse(c.CellText, out value))
						temp *= value;
				}

				success = true;
				resultText = temp.ToString(CultureInfo.InvariantCulture);
				break;
			case SLDataFieldFunctionValues.StandardDeviation:
				temp = 0D;

				foreach (var c in cells.Where(c => c.DataType == CellValues.Number))
				{
					if (c.CellText == null)
					{
						++count;
						temp += c.NumericValue;
						means.Add(c.NumericValue);
					}
					else if (double.TryParse(c.CellText, out value))
					{
						++count;
						temp += value;
						means.Add(value);
					}
				}

				if (count > 0)
				{
					mean = temp / count;
					temp = 0D;

					for (var i = 0; i < means.Count; ++i)
						temp += (mean - means[i]) * (mean - means[i]);

					temp = Math.Sqrt(temp / count);

					success = true;
					resultText = temp.ToString(CultureInfo.InvariantCulture);
				}
				else
				{
					success = false;
					resultText = SLConstants.ErrorDivisionByZero;
				}

				break;
			case SLDataFieldFunctionValues.Sum:
				temp = 0D;

				foreach (var c in cells.Where(c => c.DataType == CellValues.Number))
				{
					if (c.CellText == null)
						temp += c.NumericValue;
					else if (double.TryParse(c.CellText, out value))
						temp += value;
				}

				success = true;
				resultText = temp.ToString(CultureInfo.InvariantCulture);
				break;
			case SLDataFieldFunctionValues.Variance:
				temp = 0D;
				mean = 0D;

				foreach (var c in cells.Where(c => c.DataType == CellValues.Number))
				{
					if (c.CellText == null)
					{
						++count;
						mean += c.NumericValue;
						temp += c.NumericValue * c.NumericValue;
					}
					else if (double.TryParse(c.CellText, out value))
					{
						++count;
						mean += value;
						temp += (value * value);
					}
				}

				if (count <= 1)
				{
					success = false;
					resultText = SLConstants.ErrorDivisionByZero;
				}
				else
				{
					success = true;
					--count;
					temp = (mean / count) - ((temp / count) * (temp / count));
					resultText = temp.ToString(CultureInfo.InvariantCulture);
				}

				break;
		}

		return success;
	}

	internal static int GetFunctionNumber(TotalsRowFunctionValues function)
	{
		return function switch
		{
			TotalsRowFunctionValues.Average => 101,
			TotalsRowFunctionValues.Count => 103,
			TotalsRowFunctionValues.CountNumbers => 102,
			TotalsRowFunctionValues.Maximum => 104,
			TotalsRowFunctionValues.Minimum => 105,
			TotalsRowFunctionValues.StandardDeviation => 107,
			TotalsRowFunctionValues.Sum => 109,
			TotalsRowFunctionValues.Variance => 110,
			_ => 0
		};
	}
}
