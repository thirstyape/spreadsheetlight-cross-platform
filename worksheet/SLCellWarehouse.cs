using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpreadsheetLight
{
    internal class SLCellWarehouse
    {
        internal Dictionary<int, Dictionary<int, SLCell>> Cells { get; set; }

        internal SLCellWarehouse()
        {
            this.Cells = new Dictionary<int, Dictionary<int, SLCell>>();
        }

        internal void SetValue(int RowIndex, int ColumnIndex, SLCell Cell)
        {
            if (!this.Cells.ContainsKey(RowIndex))
            {
                this.Cells.Add(RowIndex, new Dictionary<int, SLCell>());
            }

            this.Cells[RowIndex][ColumnIndex] = Cell.Clone();
        }

        // we are not going to implement this because there's no need
        // If the cell already exists, I want that. If not, then a new SLCell.
        // The calling function would need to know which of the 2 situations above happens.
        // Returning an SLCell without knowing which would then be useless.
        //public SLCell GetValue(int RowIndex, int ColumnIndex)
        //{
        //    SLCell result = new SLCell();
        //    if (this.Cells.ContainsKey(RowIndex) && this.Cells[RowIndex].ContainsKey(ColumnIndex))
        //    {
        //        result = this.Cells[RowIndex][ColumnIndex].Clone();
        //    }

        //    return result;
        //}

        internal bool Exists(int RowIndex, int ColumnIndex)
        {
            bool result = false;
            if (this.Cells.ContainsKey(RowIndex) && this.Cells[RowIndex].ContainsKey(ColumnIndex))
            {
                result = true;
            }
            return result;
        }

        internal bool Remove(int RowIndex, int ColumnIndex)
        {
            bool result = false;
            if (this.Cells.ContainsKey(RowIndex) && this.Cells[RowIndex].ContainsKey(ColumnIndex))
            {
                result = this.Cells[RowIndex].Remove(ColumnIndex);
            }
            return result;
        }

        internal void Clear()
        {
            this.Cells.Clear();
        }
    }
}
