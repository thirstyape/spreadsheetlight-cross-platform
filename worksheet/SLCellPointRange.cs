using System;
using System.Collections.Generic;

namespace SpreadsheetLight
{
    /// <summary>
    /// This represents a cell reference range in numeric index form.
    /// </summary>
    public struct SLCellPointRange
    {
        /// <summary>
        /// Start row index.
        /// </summary>
        public int StartRowIndex;

        /// <summary>
        /// Start column index.
        /// </summary>
        public int StartColumnIndex;

        /// <summary>
        /// End row index.
        /// </summary>
        public int EndRowIndex;

        /// <summary>
        /// End column index.
        /// </summary>
        public int EndColumnIndex;

        /// <summary>
        /// Initializes an instance of SLCellPointRange.
        /// </summary>
        /// <param name="StartRowIndex">The start row index.</param>
        /// <param name="StartColumnIndex">The start column index.</param>
        /// <param name="EndRowIndex">The end row index.</param>
        /// <param name="EndColumnIndex">The end column index.</param>
        public SLCellPointRange(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            this.StartRowIndex = StartRowIndex;
            this.StartColumnIndex = StartColumnIndex;
            this.EndRowIndex = EndRowIndex;
            this.EndColumnIndex = EndColumnIndex;
        }
    }

    internal class SLCellPointRangeComparer : IComparer<SLCellPointRange>
    {
        public int Compare(SLCellPointRange pt1, SLCellPointRange pt2)
        {
            if (pt1.StartRowIndex < pt2.StartRowIndex)
            {
                return -1;
            }
            else if (pt1.StartRowIndex > pt2.StartRowIndex)
            {
                return 1;
            }
            else
            {
                if (pt1.StartColumnIndex < pt2.StartColumnIndex)
                {
                    return -1;
                }
                else if (pt1.StartColumnIndex > pt2.StartColumnIndex)
                {
                    return 1;
                }
                else
                {
                    if (pt1.EndRowIndex < pt2.EndRowIndex)
                    {
                        return -1;
                    }
                    else if (pt1.EndRowIndex > pt2.EndRowIndex)
                    {
                        return 1;
                    }
                    else
                    {
                        return pt1.EndColumnIndex.CompareTo(pt2.EndColumnIndex);
                    }
                }
            }
        }
    }
}
