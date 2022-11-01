using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLSheetView
    {
        // Part of the set of properties is also in SLPageSettings.
        // Short answer is that it's easier for the developer to access all page-like
        // properties in one class. That's why this class isn't exposed.
        // So remember to sync whenever relevant.

        internal bool HasPane
        {
            get
            {
                return this.Pane.HorizontalSplit != 0 || this.Pane.VerticalSplit != 0
                    || this.Pane.TopLeftCell != null || this.Pane.ActivePane != PaneValues.TopLeft
                    || this.Pane.State != PaneStateValues.Split;
            }
        }

        internal SLPane Pane { get; set; }

        internal List<SLSelection> Selections { get; set; }
        internal List<PivotSelection> PivotSelections { get; set; }

        internal bool WindowProtection { get; set; }

        // it appears that this doesn't change the showing/hiding of formula bar in Excel.
        // Excel has an option that does this, but is outside of the control of the Open XML
        // document. So it doesn't matter if you're using Open XML SDK or just rendering
        // XML tags. You're not going to control it.
        // Why? I don't know. Ask the Microsoft developer responsible.
        // It appears that if you change that option (within Excel), *all* Excel spreadsheets either
        // show or hide the formula bar for all worksheets.
        // Why this option is even here is beyond me...
        internal bool ShowFormulas { get; set; }

        internal bool ShowGridLines { get; set; }
        internal bool ShowRowColHeaders { get; set; }
        internal bool ShowZeros { get; set; }
        internal bool RightToLeft { get; set; }
        internal bool TabSelected { get; set; }
        internal bool ShowRuler { get; set; }
        internal bool ShowOutlineSymbols { get; set; }
        internal bool DefaultGridColor { get; set; }
        internal bool ShowWhiteSpace { get; set; }
        internal SheetViewValues View { get; set; }
        internal string TopLeftCell { get; set; }
        internal uint ColorId { get; set; }

        internal uint iZoomScale;
        internal uint ZoomScale
        {
            get { return iZoomScale; }
            set
            {
                iZoomScale = value;
                if (iZoomScale < 10) iZoomScale = 10;
                if (iZoomScale > 400) iZoomScale = 400;
            }
        }

        internal uint iZoomScaleNormal;
        internal uint ZoomScaleNormal
        {
            get { return iZoomScaleNormal; }
            set
            {
                iZoomScaleNormal = value;
                if (iZoomScaleNormal < 10) iZoomScaleNormal = 10;
                if (iZoomScaleNormal > 400) iZoomScaleNormal = 400;
            }
        }

        internal uint iZoomScaleSheetLayoutView;
        internal uint ZoomScaleSheetLayoutView
        {
            get { return iZoomScaleSheetLayoutView; }
            set
            {
                iZoomScaleSheetLayoutView = value;
                if (iZoomScaleSheetLayoutView < 10) iZoomScaleSheetLayoutView = 10;
                if (iZoomScaleSheetLayoutView > 400) iZoomScaleSheetLayoutView = 400;
            }
        }

        internal uint iZoomScalePageLayoutView;
        internal uint ZoomScalePageLayoutView
        {
            get { return iZoomScalePageLayoutView; }
            set
            {
                iZoomScalePageLayoutView = value;
                if (iZoomScalePageLayoutView < 10) iZoomScalePageLayoutView = 10;
                if (iZoomScalePageLayoutView > 400) iZoomScalePageLayoutView = 400;
            }
        }

        internal uint WorkbookViewId { get; set; }

        internal SLSheetView()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Pane = new SLPane();

            this.Selections = new List<SLSelection>();
            this.PivotSelections = new List<PivotSelection>();

            this.WindowProtection = false;
            this.ShowFormulas = false;
            this.ShowGridLines = true;
            this.ShowRowColHeaders = true;
            this.ShowZeros = true;
            this.RightToLeft = false;
            this.TabSelected = false;
            this.ShowRuler = true;
            this.ShowOutlineSymbols = true;
            this.DefaultGridColor = true;
            this.ShowWhiteSpace = true;
            this.View = SheetViewValues.Normal;
            this.TopLeftCell = string.Empty;
            this.ColorId = 64;
            this.iZoomScale = 100;
            this.iZoomScaleNormal = 0;
            this.iZoomScaleSheetLayoutView = 0;
            this.iZoomScalePageLayoutView = 0;
            this.WorkbookViewId = 0;
        }

        internal void FromSheetView(SheetView sv)
        {
            this.SetAllNull();

            if (sv.WindowProtection != null) this.WindowProtection = sv.WindowProtection.Value;
            if (sv.ShowFormulas != null) this.ShowFormulas = sv.ShowFormulas.Value;
            if (sv.ShowGridLines != null) this.ShowGridLines = sv.ShowGridLines.Value;
            if (sv.ShowRowColHeaders != null) this.ShowRowColHeaders = sv.ShowRowColHeaders.Value;
            if (sv.ShowZeros != null) this.ShowZeros = sv.ShowZeros.Value;
            if (sv.RightToLeft != null) this.RightToLeft = sv.RightToLeft.Value;
            if (sv.TabSelected != null) this.TabSelected = sv.TabSelected.Value;
            if (sv.ShowRuler != null) this.ShowRuler = sv.ShowRuler.Value;
            if (sv.ShowOutlineSymbols != null) this.ShowOutlineSymbols = sv.ShowOutlineSymbols.Value;
            if (sv.DefaultGridColor != null) this.DefaultGridColor = sv.DefaultGridColor.Value;
            if (sv.ShowWhiteSpace != null) this.ShowWhiteSpace = sv.ShowWhiteSpace.Value;
            if (sv.View != null) this.View = sv.View.Value;
            if (sv.TopLeftCell != null) this.TopLeftCell = sv.TopLeftCell.Value;
            if (sv.ColorId != null) this.ColorId = sv.ColorId.Value;
            if (sv.ZoomScale != null) this.ZoomScale = sv.ZoomScale.Value;
            if (sv.ZoomScaleNormal != null) this.ZoomScaleNormal = sv.ZoomScaleNormal.Value;
            if (sv.ZoomScaleSheetLayoutView != null) this.ZoomScaleSheetLayoutView = sv.ZoomScaleSheetLayoutView.Value;
            if (sv.ZoomScalePageLayoutView != null) this.ZoomScalePageLayoutView = sv.ZoomScalePageLayoutView.Value;

            // required attribute but we'll use 0 as the default in case something terrible happens.
            if (sv.WorkbookViewId != null) this.WorkbookViewId = sv.WorkbookViewId.Value;
            else this.WorkbookViewId = 0;

            using (OpenXmlReader oxr = OpenXmlReader.Create(sv))
            {
                SLSelection sel;
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Pane))
                    {
                        this.Pane = new SLPane();
                        this.Pane.FromPane((Pane)oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(Selection))
                    {
                        sel = new SLSelection();
                        sel.FromSelection((Selection)oxr.LoadCurrentElement());
                        this.Selections.Add(sel);
                    }
                    else if (oxr.ElementType == typeof(PivotSelection))
                    {
                        this.PivotSelections.Add((PivotSelection)oxr.LoadCurrentElement().CloneNode(true));
                    }
                }
            }
        }

        internal SheetView ToSheetView()
        {
            SheetView sv = new SheetView();
            if (this.WindowProtection != false) sv.WindowProtection = this.WindowProtection;
            if (this.ShowFormulas != false) sv.ShowFormulas = this.ShowFormulas;
            if (this.ShowGridLines != true) sv.ShowGridLines = this.ShowGridLines;
            if (this.ShowRowColHeaders != true) sv.ShowRowColHeaders = this.ShowRowColHeaders;
            if (this.ShowZeros != true) sv.ShowZeros = this.ShowZeros;
            if (this.RightToLeft != false) sv.RightToLeft = this.RightToLeft;
            if (this.TabSelected != false) sv.TabSelected = this.TabSelected;
            if (this.ShowRuler != true) sv.ShowRuler = this.ShowRuler;
            if (this.ShowOutlineSymbols != true) sv.ShowOutlineSymbols = this.ShowOutlineSymbols;
            if (this.DefaultGridColor != true) sv.DefaultGridColor = this.DefaultGridColor;
            if (this.ShowWhiteSpace != true) sv.ShowWhiteSpace = this.ShowWhiteSpace;
            if (this.View != SheetViewValues.Normal) sv.View = this.View;
            if (this.TopLeftCell != null && this.TopLeftCell.Length > 0) sv.TopLeftCell = this.TopLeftCell;
            if (this.ColorId != 64) sv.ColorId = this.ColorId;
            if (this.ZoomScale != 100) sv.ZoomScale = this.ZoomScale;
            if (this.ZoomScaleNormal != 0) sv.ZoomScaleNormal = this.ZoomScaleNormal;
            if (this.ZoomScaleSheetLayoutView != 0) sv.ZoomScaleSheetLayoutView = this.ZoomScaleSheetLayoutView;
            if (this.ZoomScalePageLayoutView != 0) sv.ZoomScalePageLayoutView = this.ZoomScalePageLayoutView;
            sv.WorkbookViewId = this.WorkbookViewId;

            if (HasPane)
            {
                sv.Append(this.Pane.ToPane());
            }

            foreach (SLSelection sel in this.Selections)
            {
                sv.Append(sel.ToSelection());
            }

            foreach (PivotSelection psel in this.PivotSelections)
            {
                sv.Append((PivotSelection)psel.CloneNode(true));
            }

            return sv;
        }

        internal SLSheetView Clone()
        {
            SLSheetView sv = new SLSheetView();
            sv.Pane = this.Pane.Clone();

            sv.Selections = new List<SLSelection>();
            foreach (SLSelection sel in this.Selections)
            {
                sv.Selections.Add(sel.Clone());
            }

            sv.PivotSelections = new List<PivotSelection>();
            foreach (PivotSelection psel in this.PivotSelections)
            {
                sv.PivotSelections.Add((PivotSelection)psel.CloneNode(true));
            }

            sv.WindowProtection = this.WindowProtection;
            sv.ShowFormulas = this.ShowFormulas;
            sv.ShowGridLines = this.ShowGridLines;
            sv.ShowRowColHeaders = this.ShowRowColHeaders;
            sv.ShowZeros = this.ShowZeros;
            sv.RightToLeft = this.RightToLeft;
            sv.TabSelected = this.TabSelected;
            sv.ShowRuler = this.ShowRuler;
            sv.ShowOutlineSymbols = this.ShowOutlineSymbols;
            sv.DefaultGridColor = this.DefaultGridColor;
            sv.ShowWhiteSpace = this.ShowWhiteSpace;
            sv.View = this.View;
            sv.TopLeftCell = this.TopLeftCell;
            sv.ColorId = this.ColorId;
            sv.iZoomScale = this.iZoomScale;
            sv.iZoomScaleNormal = this.iZoomScaleNormal;
            sv.iZoomScaleSheetLayoutView = this.iZoomScaleSheetLayoutView;
            sv.iZoomScalePageLayoutView = this.iZoomScalePageLayoutView;
            sv.WorkbookViewId = this.WorkbookViewId;

            return sv;
        }

        internal static string GetSheetViewValuesAttribute(SheetViewValues svv)
        {
            string result = "normal";
            switch (svv)
            {
                case SheetViewValues.Normal:
                    result = "normal";
                    break;
                case SheetViewValues.PageBreakPreview:
                    result = "pageBreakPreview";
                    break;
                case SheetViewValues.PageLayout:
                    result = "pageLayout";
                    break;
            }

            return result;
        }
    }
}
