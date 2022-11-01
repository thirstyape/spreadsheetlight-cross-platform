using System;
using System.Collections.Generic;

namespace SpreadsheetLight
{
    /// <summary>
    /// This represents a cell reference in numeric index form.
    /// </summary>
    public struct SLCellPoint
    {
        /// <summary>
        /// Row index.
        /// </summary>
        public int RowIndex;

        /// <summary>
        /// Column index.
        /// </summary>
        public int ColumnIndex;

        // is this even useful as a public constructor? Whatever...
        /// <summary>
        /// Initializes an instance of SLCellPoint.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        public SLCellPoint(int RowIndex, int ColumnIndex)
        {
            this.RowIndex = RowIndex;
            this.ColumnIndex = ColumnIndex;
        }

        /// <summary>
        /// Returns the hash code for this instance.
        /// </summary>
        /// <returns>The hash code.</returns>
        public override int GetHashCode()
        {
            //http://stackoverflow.com/questions/263400/what-is-the-best-algorithm-for-an-overridden-system-object-gethashcode?lq=1
            //http://stackoverflow.com/questions/1646807/quick-and-simple-hash-code-combinations/1646913#1646913
            // Thanks Jon Skeet!
            // The "unchecked" part wraps calculations when overflow.
            // There's a small issue with speed though...
            // Apparently, the hashing algorithm affects the speed of how the Dictionary data structure
            // stores the cells, with this cell point structure as the key.

            // OPTION 1:
            // This works better with a large number of rows and small number of columns.
            // Say 50000 rows by 50 columns.
            // Which is probably most of the time, judging from people asking
            // "Can your library support millions of cells?"
            //unchecked
            //{
            //    int hash = 17;
            //    hash = hash * 31 + RowIndex.GetHashCode();
            //    hash = hash * 31 + ColumnIndex.GetHashCode();
            //    return hash;
            //}

            // OPTION 2:
            // This works better with a large number of columns and small number of rows.
            // Say 156 rows by 16000 columns.
            //unchecked
            //{
            //    int hash = 17;
            //    hash = hash * 31 + ColumnIndex.GetHashCode();
            //    hash = hash * 31 + RowIndex.GetHashCode();
            //    return hash;
            //}

            // OPTION 3:
            // This option appears to be only slightly slower than the best of the 2 hashing algorithms
            // above. However, it works consistently whether you have large numbers of rows or
            // large numbers of columns. So I'm going to use this.
            return string.Format("{0}{1}", RowIndex.ToString("d7", System.Globalization.CultureInfo.InvariantCulture), ColumnIndex.ToString("d5", System.Globalization.CultureInfo.InvariantCulture)).GetHashCode();

            // For academic interest, here's a sample of the empirical data. Cells are set with a
            // mixture of integers, floating point values and strings. String data are in the form
            // of "R{0}C{1}". So row 5 column 7 is "R5C7". This ensures no string is ever duplicated.
            // This is to "bulk up" the shared string table.
            // Data is in this ratio: integer 40%, floating point 40%, string 20%
            // Well, approximately anyway...
            
            // For 156 rows by 16000 columns by 3 worksheets:
            // Option 1: 2 minutes 44 seconds
            // Option 2: 1 minute 2 seconds
            // Option 3: 1 minute 7 seconds

            // For 50000 rows by 50 columns by 3 worksheets:
            // Option 1: 59 seconds
            // Option 2: 1 minute 29 seconds
            // Option 3: 1 minute 7 seconds

            // And if the 3rd option performs badly in terms of hashing collisions, try option 2.
            // I don't know why having ColumnIndex before RowIndex makes it faster.
            // My guess is that it makes the final hash "smaller" so it doesn't wrap around (column
            // indices are smaller in magnitude than row indices).
            // It appears from empirical experiments that the hash code of integer 8000 is also 8000.
            // So it seems as if the hash code of a 32-bit integer is itself.
            // DISCLAIMER: But I'm not familiar with hashing algorithms or with GetHashCode().

            // And yes, I realise that the amount of comments in this section completely overshadow
            // the amount of actual working code.
        }
    }

    internal class SLCellReferencePointComparer : IComparer<SLCellPoint>
    {
        public int Compare(SLCellPoint pt1, SLCellPoint pt2)
        {
            if (pt1.RowIndex < pt2.RowIndex)
            {
                return -1;
            }
            else if (pt1.RowIndex > pt2.RowIndex)
            {
                return 1;
            }
            else
            {
                return pt1.ColumnIndex.CompareTo(pt2.ColumnIndex);
            }
        }
    }
}
