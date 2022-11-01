using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLSelection
    {
        internal PaneValues Pane { get; set; }
        internal string ActiveCell { get; set; }
        internal uint ActiveCellId { get; set; }
        internal List<SLCellPointRange> SequenceOfReferences { get; set; }

        internal SLSelection()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Pane = PaneValues.TopLeft;
            this.ActiveCell = string.Empty;
            this.ActiveCellId = 0;
            this.SequenceOfReferences = new List<SLCellPointRange>();
        }

        internal void FromSelection(Selection s)
        {
            this.SetAllNull();

            if (s.Pane != null) this.Pane = s.Pane.Value;
            if (s.ActiveCell != null) this.ActiveCell = s.ActiveCell.Value;
            if (s.ActiveCellId != null) this.ActiveCellId = s.ActiveCellId.Value;
            if (s.SequenceOfReferences != null) this.SequenceOfReferences = SLTool.TranslateSeqRefToCellPointRange(s.SequenceOfReferences);
        }

        internal Selection ToSelection()
        {
            Selection s = new Selection();
            if (this.Pane != PaneValues.TopLeft) s.Pane = this.Pane;

            if (this.ActiveCell.Length > 0 && !this.ActiveCell.Equals("A1", StringComparison.OrdinalIgnoreCase))
            {
                s.ActiveCell = this.ActiveCell;
            }
            
            if (this.ActiveCellId != 0) s.ActiveCellId = this.ActiveCellId;

            if (this.SequenceOfReferences.Count > 0)
            {
                if (this.SequenceOfReferences.Count == 1)
                {
                    // not equal to A1
                    if (this.SequenceOfReferences[0].StartRowIndex != 1
                        || this.SequenceOfReferences[0].StartColumnIndex != 1
                        || this.SequenceOfReferences[0].EndRowIndex != 1
                        || this.SequenceOfReferences[0].EndColumnIndex != 1)
                    {
                        s.SequenceOfReferences = SLTool.TranslateCellPointRangeToSeqRef(this.SequenceOfReferences);
                    }
                }
                else
                {
                    s.SequenceOfReferences = SLTool.TranslateCellPointRangeToSeqRef(this.SequenceOfReferences);
                }
            }

            return s;
        }

        internal SLSelection Clone()
        {
            SLSelection s = new SLSelection();
            s.Pane = this.Pane;
            s.ActiveCell = this.ActiveCell;
            s.ActiveCellId = this.ActiveCellId;

            s.SequenceOfReferences = new List<SLCellPointRange>();
            foreach (SLCellPointRange pt in this.SequenceOfReferences)
            {
                s.SequenceOfReferences.Add(new SLCellPointRange(pt.StartRowIndex, pt.StartColumnIndex, pt.EndRowIndex, pt.EndColumnIndex));
            }

            return s;
        }
    }
}
