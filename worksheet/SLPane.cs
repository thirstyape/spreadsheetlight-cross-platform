using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLPane
    {
        internal double HorizontalSplit { get; set; }
        internal double VerticalSplit { get; set; }
        internal string TopLeftCell { get; set; }
        internal PaneValues ActivePane { get; set; }
        internal PaneStateValues State { get; set; }

        internal SLPane()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.HorizontalSplit = 0;
            this.VerticalSplit = 0;
            this.TopLeftCell = null;
            this.ActivePane = PaneValues.TopLeft;
            this.State = PaneStateValues.Split;
        }

        internal void FromPane(Pane p)
        {
            this.SetAllNull();

            if (p.HorizontalSplit != null) this.HorizontalSplit = p.HorizontalSplit.Value;
            if (p.VerticalSplit != null) this.VerticalSplit = p.VerticalSplit.Value;
            if (p.TopLeftCell != null) this.TopLeftCell = p.TopLeftCell.Value;
            if (p.ActivePane != null) this.ActivePane = p.ActivePane.Value;
            if (p.State != null) this.State = p.State.Value;
        }

        internal Pane ToPane()
        {
            Pane p = new Pane();
            if (this.HorizontalSplit != 0) p.HorizontalSplit = this.HorizontalSplit;
            if (this.VerticalSplit != 0) p.VerticalSplit = this.VerticalSplit;
            if (this.TopLeftCell != null && this.TopLeftCell.Length > 0) p.TopLeftCell = this.TopLeftCell;
            if (this.ActivePane != PaneValues.TopLeft) p.ActivePane = this.ActivePane;
            if (this.State != PaneStateValues.Split) p.State = this.State;

            return p;
        }

        internal SLPane Clone()
        {
            SLPane p = new SLPane();
            p.HorizontalSplit = this.HorizontalSplit;
            p.VerticalSplit = this.VerticalSplit;
            p.TopLeftCell = this.TopLeftCell;
            p.ActivePane = this.ActivePane;
            p.State = this.State;

            return p;
        }

        internal static string GetPaneValuesAttribute(PaneValues pv)
        {
            string result = "topLeft";
            switch (pv)
            {
                case PaneValues.BottomLeft:
                    result = "bottomLeft";
                    break;
                case PaneValues.BottomRight:
                    result = "bottomRight";
                    break;
                case PaneValues.TopLeft:
                    result = "topLeft";
                    break;
                case PaneValues.TopRight:
                    result = "topRight";
                    break;
            }

            return result;
        }

        internal static string GetPaneStateValuesAttribute(PaneStateValues psv)
        {
            string result = "split";
            switch (psv)
            {
                case PaneStateValues.Frozen:
                    result = "frozen";
                    break;
                case PaneStateValues.FrozenSplit:
                    result = "frozenSplit";
                    break;
                case PaneStateValues.Split:
                    result = "split";
                    break;
            }

            return result;
        }
    }
}
