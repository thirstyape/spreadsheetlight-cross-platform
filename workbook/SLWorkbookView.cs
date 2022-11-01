using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLWorkbookView
    {
        internal VisibilityValues Visibility { get; set; }
        internal bool Minimized { get; set; }
        internal bool ShowHorizontalScroll { get; set; }
        internal bool ShowVerticalScroll { get; set; }
        internal bool ShowSheetTabs { get; set; }
        internal int? XWindow { get; set; }
        internal int? YWindow { get; set; }
        internal uint? WindowWidth { get; set; }
        internal uint? WindowHeight { get; set; }
        internal uint TabRatio { get; set; }
        internal uint FirstSheet { get; set; }
        internal uint ActiveTab { get; set; }
        internal bool AutoFilterDateGrouping { get; set; }

        internal SLWorkbookView()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Visibility = VisibilityValues.Visible;
            this.Minimized = false;
            this.ShowHorizontalScroll = true;
            this.ShowVerticalScroll = true;
            this.ShowSheetTabs = true;
            this.XWindow = null;
            this.YWindow = null;
            this.WindowWidth = null;
            this.WindowHeight = null;
            this.TabRatio = 600;
            this.FirstSheet = 0;
            this.ActiveTab = 0;
            this.AutoFilterDateGrouping = true;
        }

        internal void FromWorkbookView(WorkbookView wv)
        {
            this.SetAllNull();

            if (wv.Visibility != null) this.Visibility = wv.Visibility.Value;
            if (wv.Minimized != null) this.Minimized = wv.Minimized.Value;
            if (wv.ShowHorizontalScroll != null) this.ShowHorizontalScroll = wv.ShowHorizontalScroll.Value;
            if (wv.ShowVerticalScroll != null) this.ShowVerticalScroll = wv.ShowVerticalScroll.Value;
            if (wv.ShowSheetTabs != null) this.ShowSheetTabs = wv.ShowSheetTabs.Value;
            if (wv.XWindow != null) this.XWindow = wv.XWindow.Value;
            if (wv.YWindow != null) this.YWindow = wv.YWindow.Value;
            if (wv.WindowWidth != null) this.WindowWidth = wv.WindowWidth.Value;
            if (wv.WindowHeight != null) this.WindowHeight = wv.WindowHeight.Value;
            if (wv.TabRatio != null) this.TabRatio = wv.TabRatio.Value;
            if (wv.FirstSheet != null) this.FirstSheet = wv.FirstSheet.Value;
            if (wv.ActiveTab != null) this.ActiveTab = wv.ActiveTab.Value;
            if (wv.AutoFilterDateGrouping != null) this.AutoFilterDateGrouping = wv.AutoFilterDateGrouping.Value;
        }

        internal WorkbookView ToWorkbookView()
        {
            WorkbookView wv = new WorkbookView();
            if (this.Visibility != VisibilityValues.Visible) wv.Visibility = this.Visibility;
            if (this.Minimized) wv.Minimized = this.Minimized;
            if (!this.ShowHorizontalScroll) wv.ShowHorizontalScroll = this.ShowHorizontalScroll;
            if (!this.ShowVerticalScroll) wv.ShowVerticalScroll = this.ShowVerticalScroll;
            if (!this.ShowSheetTabs) wv.ShowSheetTabs = this.ShowSheetTabs;
            if (this.XWindow != null) wv.XWindow = this.XWindow.Value;
            if (this.YWindow != null) wv.YWindow = this.YWindow.Value;
            if (this.WindowWidth != null) wv.WindowWidth = this.WindowWidth.Value;
            if (this.WindowHeight != null) wv.WindowHeight = this.WindowHeight.Value;
            if (this.TabRatio != 600) wv.TabRatio = this.TabRatio;
            if (this.FirstSheet != 0) wv.FirstSheet = this.FirstSheet;
            if (this.ActiveTab != 0) wv.ActiveTab = this.ActiveTab;
            if (!this.AutoFilterDateGrouping) wv.AutoFilterDateGrouping = this.AutoFilterDateGrouping;

            return wv;
        }

        internal SLWorkbookView Clone()
        {
            SLWorkbookView wv = new SLWorkbookView();
            wv.Visibility = this.Visibility;
            wv.Minimized = this.Minimized;
            wv.ShowHorizontalScroll = this.ShowHorizontalScroll;
            wv.ShowVerticalScroll = this.ShowVerticalScroll;
            wv.ShowSheetTabs = this.ShowSheetTabs;
            wv.XWindow = this.XWindow;
            wv.YWindow = this.YWindow;
            wv.WindowWidth = this.WindowWidth;
            wv.WindowHeight = this.WindowHeight;
            wv.TabRatio = this.TabRatio;
            wv.FirstSheet = this.FirstSheet;
            wv.ActiveTab = this.ActiveTab;
            wv.AutoFilterDateGrouping = this.AutoFilterDateGrouping;

            return wv;
        }
    }
}
