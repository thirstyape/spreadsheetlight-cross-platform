using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLHyperlink
    {
        internal bool IsExternal;
        internal string HyperlinkUri;
        internal UriKind HyperlinkUriKind;
        internal bool IsNew;

        internal SLCellPointRange Reference { get; set; }
        internal string Id { get; set; }
        internal string Location { get; set; }
        internal string ToolTip { get; set; }
        internal string Display { get; set; }

        internal SLHyperlink()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.IsExternal = false;
            this.HyperlinkUri = string.Empty;
            this.HyperlinkUriKind = UriKind.RelativeOrAbsolute;
            this.IsNew = true;

            this.Reference = new SLCellPointRange();
            this.Id = null;
            this.Location = null;
            this.ToolTip = null;
            this.Display = null;
        }

        internal void FromHyperlink(Hyperlink hl)
        {
            this.SetAllNull();

            this.IsNew = false;

            if (hl.Reference != null) this.Reference = SLTool.TranslateReferenceToCellPointRange(hl.Reference.Value);

            if (hl.Id != null)
            {
                // At least I think if there's a relationship ID, it's an external link.
                this.IsExternal = true;
                this.Id = hl.Id.Value;
            }

            if (hl.Location != null) this.Location = hl.Location.Value;
            if (hl.Tooltip != null) this.ToolTip = hl.Tooltip.Value;
            if (hl.Display != null) this.Display = hl.Display.Value;
        }

        internal Hyperlink ToHyperlink()
        {
            Hyperlink hl = new Hyperlink();
            hl.Reference = SLTool.ToCellRange(this.Reference.StartRowIndex, this.Reference.StartColumnIndex, this.Reference.EndRowIndex, this.Reference.EndColumnIndex);
            if (this.Id != null && this.Id.Length > 0) hl.Id = this.Id;
            if (this.Location != null && this.Location.Length > 0) hl.Location = this.Location;
            if (this.ToolTip != null && this.ToolTip.Length > 0) hl.Tooltip = this.ToolTip;
            if (this.Display != null && this.Display.Length > 0) hl.Display = this.Display;

            return hl;
        }

        internal SLHyperlink Clone()
        {
            SLHyperlink hl = new SLHyperlink();
            hl.IsExternal = this.IsExternal;
            hl.HyperlinkUri = this.HyperlinkUri;
            hl.HyperlinkUriKind = this.HyperlinkUriKind;
            hl.IsNew = this.IsNew;
            hl.Reference = new SLCellPointRange(this.Reference.StartRowIndex, this.Reference.StartColumnIndex, this.Reference.EndRowIndex, this.Reference.EndColumnIndex);
            hl.Id = this.Id;
            hl.Location = this.Location;
            hl.ToolTip = this.ToolTip;
            hl.Display = this.Display;

            return hl;
        }
    }
}
