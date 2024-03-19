using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight;

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

        if (pv == PaneValues.BottomLeft)
            result = "bottomLeft";
        else if (pv == PaneValues.BottomRight)
            result = "bottomRight";
        else if (pv == PaneValues.TopLeft)
            result = "topLeft";
        else if (pv == PaneValues.TopRight)
            result = "topRight";

        return result;
    }

    internal static string GetPaneStateValuesAttribute(PaneStateValues psv)
    {
        string result = "split";

        if (psv == PaneStateValues.Frozen)
            result = "frozen";
        else if (psv == PaneStateValues.FrozenSplit)
            result = "frozenSplit";
        else if (psv == PaneStateValues.Split)
            result = "split";

        return result;
    }
}
