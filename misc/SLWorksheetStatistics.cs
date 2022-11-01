using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpreadsheetLight
{
    /// <summary>
    /// Statistical information about a worksheet.
    /// </summary>
    public class SLWorksheetStatistics
    {
        internal int iStartRowIndex;
        /// <summary>
        /// Index of the first row used. This includes empty rows but might be styled. This returns -1 if the worksheet is empty (but check for negative values instead of -1 just in case). This is read-only.
        /// </summary>
        public int StartRowIndex { get { return iStartRowIndex; } }

        internal int iStartColumnIndex;
        /// <summary>
        /// Index of the first column used. This includes empty columns but might be styled. This returns -1 if the worksheet is empty (but check for negative values instead of -1 just in case). This is read-only.
        /// </summary>
        public int StartColumnIndex { get { return iStartColumnIndex; } }

        internal int iEndRowIndex;
        /// <summary>
        /// Index of the last row used. This includes empty rows but might be styled. This returns -1 if the worksheet is empty (but check for negative values instead of -1 just in case). This is read-only.
        /// </summary>
        public int EndRowIndex { get { return iEndRowIndex; } }

        internal int iEndColumnIndex;
        /// <summary>
        /// Index of the last column used. This includes empty columns but might be styled. This returns -1 if the worksheet is empty (but check for negative values instead of -1 just in case). This is read-only.
        /// </summary>
        public int EndColumnIndex { get { return iEndColumnIndex; } }

        internal int iNumberOfCells;
        /// <summary>
        /// Number of cells set in the worksheet. This is read-only.
        /// </summary>
        public int NumberOfCells { get { return iNumberOfCells; } }

        internal int iNumberOfEmptyCells;
        /// <summary>
        /// Number of cells set in the worksheet that is empty. This could be that a style was set but no cell value given. This is read-only.
        /// </summary>
        public int NumberOfEmptyCells { get { return iNumberOfEmptyCells; } }

        internal int iNumberOfRows;
        /// <summary>
        /// Number of rows in the worksheet. This includes empty rows (no cells in that row but a row style was applied, or that row only has empty cells). This is read-only.
        /// </summary>
        public int NumberOfRows { get { return iNumberOfRows; } }

        internal int iNumberOfColumns;
        /// <summary>
        /// Number of columns in the worksheet. This includes empty columns (no cells in that column but a column style was applied, or that column only has empty cells). This is read-only.
        /// </summary>
        public int NumberOfColumns { get { return iNumberOfColumns; } }

        /// <summary>
        /// Initializes an instance of SLWorksheetStatistics. But it's quite useless on its own. Use GetWorksheetStatistics() of the SLDocument class.
        /// </summary>
        public SLWorksheetStatistics()
        {
            this.iStartRowIndex = -1;
            this.iStartColumnIndex = -1;
            this.iEndRowIndex = -1;
            this.iEndColumnIndex = -1;
            this.iNumberOfCells = 0;
            this.iNumberOfEmptyCells = 0;
            this.iNumberOfRows = 0;
            this.iNumberOfColumns = 0;
        }
    }
}
