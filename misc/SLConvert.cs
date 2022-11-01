using System;

namespace SpreadsheetLight
{
    // Why is this similar to SLTool? Why not make SLTool public and be done with it?
    // See comments at the top of SLTool.cs.

    /// <summary>
    /// Encapsulates methods for miscellaneous convertions.
    /// </summary>
    public class SLConvert
    {
        /// <summary>
        /// Get the column name given the column index.
        /// </summary>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The column name.</returns>
        public static string ToColumnName(int ColumnIndex)
        {
            return SLTool.ToColumnName(ColumnIndex);
        }

        /// <summary>
        /// Get the column index given a cell reference or column name.
        /// </summary>
        /// <param name="Input">A cell reference such as "A1" or column name such as "A". If the input is invalid, then -1 is returned.</param>
        /// <returns>The column index.</returns>
        public static int ToColumnIndex(string Input)
        {
            return SLTool.ToColumnIndex(Input);
        }

        /// <summary>
        /// Get the cell reference given the row and column index. For example "A1".
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The cell reference.</returns>
        public static string ToCellReference(int RowIndex, int ColumnIndex)
        {
            return SLTool.ToCellReference(string.Empty, RowIndex, ColumnIndex, false);
        }

        /// <summary>
        /// Get the cell reference given the row and column index. For example "A1" or "$A$1".
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="IsAbsolute">True for absolute reference. False for relative reference.</param>
        /// <returns>The cell reference.</returns>
        public static string ToCellReference(int RowIndex, int ColumnIndex, bool IsAbsolute)
        {
            return SLTool.ToCellReference(string.Empty, RowIndex, ColumnIndex, IsAbsolute);
        }

        /// <summary>
        /// Get the cell reference given the worksheet name, and row and column index. For example "Sheet1!A1".
        /// </summary>
        /// <param name="WorksheetName">The worksheet name.</param>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <returns>The cell reference.</returns>
        public static string ToCellReference(string WorksheetName, int RowIndex, int ColumnIndex)
        {
            return SLTool.ToCellReference(WorksheetName, RowIndex, ColumnIndex, false);
        }

        /// <summary>
        /// Get the cell reference given the worksheet name, and row and column index. For example "Sheet1!A1" or "Sheet1!$A$1".
        /// </summary>
        /// <param name="WorksheetName">The worksheet name.</param>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        /// <param name="IsAbsolute">True for absolute reference. False for relative reference.</param>
        /// <returns>The cell reference.</returns>
        public static string ToCellReference(string WorksheetName, int RowIndex, int ColumnIndex, bool IsAbsolute)
        {
            return SLTool.ToCellReference(WorksheetName, RowIndex, ColumnIndex, IsAbsolute);
        }

        /// <summary>
        /// Get the cell range reference given a corner cell and its opposite corner cell in a cell range. For example "A1:C5".
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <returns>The cell range reference.</returns>
        public static string ToCellRange(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return SLTool.ToCellRange(string.Empty, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, false);
        }

        /// <summary>
        /// Get the cell range reference given a corner cell and its opposite corner cell in a cell range. For example "A1:C5" or "$A$1:$C$5".
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="IsAbsolute">True for absolute reference. False for relative reference.</param>
        /// <returns>The cell range reference.</returns>
        public static string ToCellRange(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool IsAbsolute)
        {
            return SLTool.ToCellRange(string.Empty, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, IsAbsolute);
        }

        /// <summary>
        /// Get the cell range reference given a worksheet name, and a corner cell and its opposite corner cell in a cell range. For example "Sheet1!A1:C5".
        /// </summary>
        /// <param name="WorksheetName">The worksheet name.</param>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <returns>The cell range reference.</returns>
        public static string ToCellRange(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return SLTool.ToCellRange(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, false);
        }

        /// <summary>
        /// Get the cell range reference given a worksheet name, and a corner cell and its opposite corner cell in a cell range. For example "Sheet1!A1:C5" or "Sheet1!$A$1:$C$5".
        /// </summary>
        /// <param name="WorksheetName">The worksheet name.</param>
        /// <param name="StartRowIndex">The row index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="StartColumnIndex">The column index of the start cell of the cell range. This is typically the top-left cell.</param>
        /// <param name="EndRowIndex">The row index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="EndColumnIndex">The column index of the end cell of the cell range. This is typically the bottom-right cell.</param>
        /// <param name="IsAbsolute">True for absolute reference. False for relative reference.</param>
        /// <returns>The cell range reference.</returns>
        public static string ToCellRange(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex, bool IsAbsolute)
        {
            return SLTool.ToCellRange(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, IsAbsolute);
        }

        /// <summary>
        /// Convert a color in hexadecimal to a System.Drawing.Color structure.
        /// </summary>
        /// <param name="HexValue">The color in hexadecimal.</param>
        /// <returns>A System.Drawing.Color structure.</returns>
        public static System.Drawing.Color ToColor(string HexValue)
        {
            return SLTool.ToColor(HexValue);
        }

        /// <summary>
        /// Convert a set of HSL color values to a System.Drawing.Color structure.
        /// </summary>
        /// <param name="Hue">The hue measured in degrees ranging from 0 to 360 degrees.</param>
        /// <param name="Saturation">The saturation ranging from 0.0 to 1.0, where 0.0 is grayscale and 1.0 is the most saturated.</param>
        /// <param name="Luminance">The luminance (sometimes known as brightness) ranging from 0.0 to 1.0, where 0.0 is effectively black and 1.0 is effectively white.</param>
        /// <returns>A System.Drawing.Color structure.</returns>
        public static System.Drawing.Color ToColor(double Hue, double Saturation, double Luminance)
        {
            return SLTool.ToColor(Hue, Saturation, Luminance);
        }

        /// <summary>
        /// Converts inches to points.
        /// </summary>
        /// <param name="Data">A value measured in inches.</param>
        /// <returns>The converted value in points.</returns>
        public static double FromInchToPoint(double Data)
        {
            return Data * 72.0;
        }

        /// <summary>
        /// Converts points to inches.
        /// </summary>
        /// <param name="Data">A value measured in points.</param>
        /// <returns>The converted value in inches.</returns>
        public static double FromPointToInch(double Data)
        {
            return Data / 72.0;
        }

        /// <summary>
        /// Converts inches to centimeters.
        /// </summary>
        /// <param name="Data">A value measured in inches.</param>
        /// <returns>The converted value in centimeters.</returns>
        public static double FromInchToCentimeter(double Data)
        {
            return Data * 2.54;
        }

        /// <summary>
        /// Converts centimeters to inches.
        /// </summary>
        /// <param name="Data">A value measured in centimeters.</param>
        /// <returns>The converted value in inches.</returns>
        public static double FromCentimeterToInch(double Data)
        {
            return Data / 2.54;
        }

        /// <summary>
        /// Converts centimeters to points.
        /// </summary>
        /// <param name="Data">A value measured in centimeters.</param>
        /// <returns>The converted value in points.</returns>
        public static double FromCentimeterToPoint(double Data)
        {
            return Data * 72.0 / 2.54;
        }

        /// <summary>
        /// Converts points to centimeters.
        /// </summary>
        /// <param name="Data">A value measured in points.</param>
        /// <returns>The converted value in centimeters.</returns>
        public static double FromPointToCentimeter(double Data)
        {
            return Data * 2.54 / 72.0;
        }

        /// <summary>
        /// Converts inches to English Metric Units.
        /// </summary>
        /// <param name="Data">A value measured in inches.</param>
        /// <returns>The converted value in English Metric Units.</returns>
        public static double FromInchToEmu(double Data)
        {
            return Data * (double)SLConstants.InchToEMU;
        }

        /// <summary>
        /// Converts English Metric Units to inches.
        /// </summary>
        /// <param name="Data">A value measured in English Metric Units.</param>
        /// <returns>The converted value in inches.</returns>
        public static double FromEmuToInch(double Data)
        {
            return Data / (double)SLConstants.InchToEMU;
        }

        /// <summary>
        /// Converts points to English Metric Units.
        /// </summary>
        /// <param name="Data">A value measured in points.</param>
        /// <returns>The converted value in English Metric Units.</returns>
        public static double FromPointToEmu(double Data)
        {
            return Data * (double)SLConstants.PointToEMU;
        }

        /// <summary>
        /// Converts English Metric Units to points.
        /// </summary>
        /// <param name="Data">A value measured in English Metric Units.</param>
        /// <returns>The converted value in points.</returns>
        public static double FromEmuToPoint(double Data)
        {
            return Data / (double)SLConstants.PointToEMU;
        }

        /// <summary>
        /// Converts centimeters to English Metric Units.
        /// </summary>
        /// <param name="Data">A value measured in centimeters.</param>
        /// <returns>The converted value in English Metric Units.</returns>
        public static double FromCentimeterToEmu(double Data)
        {
            return Data * (double)SLConstants.CentimeterToEMU;
        }

        /// <summary>
        /// Converts English Metric Units to centimeters.
        /// </summary>
        /// <param name="Data">A value measured in English Metric Units.</param>
        /// <returns>The converted value in centimeters.</returns>
        public static double FromEmuToCentimeter(double Data)
        {
            return Data / (double)SLConstants.CentimeterToEMU;
        }
    }
}
