using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLSheetProperties
    {
        internal bool HasSheetProperties
        {
            get
            {
                return this.HasTabColor || this.ApplyStyles || !this.SummaryBelow
                    || !this.SummaryRight || !this.ShowOutlineSymbols
                    || !this.AutoPageBreaks || this.FitToPage
                    || this.SyncHorizontal || this.SyncVertical || this.SyncReference.Length > 0
                    || this.TransitionEvaluation || this.TransitionEntry
                    || !this.Published || this.CodeName.Length > 0 || this.FilterMode
                    || !this.EnableFormatConditionsCalculation;
            }
        }

        internal bool HasChartSheetProperties
        {
            get
            {
                return this.HasTabColor || !this.Published || this.CodeName.Length > 0;
            }
        }

        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        internal bool HasTabColor;
        internal SLColor clrTabColor;
        internal System.Drawing.Color TabColor
        {
            get { return clrTabColor.Color; }
            set
            {
                clrTabColor.Color = value;
                HasTabColor = (clrTabColor.Color.IsEmpty) ? false : true;
            }
        }

        internal bool ApplyStyles { get; set; }
        internal bool SummaryBelow { get; set; }
        internal bool SummaryRight { get; set; }
        internal bool ShowOutlineSymbols { get; set; }

        internal bool AutoPageBreaks { get; set; }
        internal bool FitToPage { get; set; }

        internal bool SyncHorizontal { get; set; }
        internal bool SyncVertical { get; set; }
        internal string SyncReference { get; set; }
        internal bool TransitionEvaluation { get; set; }
        internal bool TransitionEntry { get; set; }
        internal bool Published { get; set; }
        internal string CodeName { get; set; }
        internal bool FilterMode { get; set; }
        internal bool EnableFormatConditionsCalculation { get; set; }

        internal SLSheetProperties(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }

            this.listIndexedColors = new List<System.Drawing.Color>();
            for (i = 0; i < IndexedColors.Count; ++i)
            {
                this.listIndexedColors.Add(IndexedColors[i]);
            }

            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.clrTabColor = new SLColor(this.listThemeColors, this.listIndexedColors);
            this.HasTabColor = false;

            this.ApplyStyles = false;
            this.SummaryBelow = true;
            this.SummaryRight = true;
            this.ShowOutlineSymbols = true;

            this.AutoPageBreaks = true;
            this.FitToPage = false;

            this.SyncHorizontal = false;
            this.SyncVertical = false;
            this.SyncReference = string.Empty;
            this.TransitionEvaluation = false;
            this.TransitionEntry = false;
            this.Published = true;
            this.CodeName = string.Empty;
            this.FilterMode = false;
            this.EnableFormatConditionsCalculation = true;
        }

        internal void FromSheetProperties(SheetProperties sp)
        {
            this.SetAllNull();
            if (sp.TabColor != null)
            {
                if (sp.TabColor.Indexed != null || sp.TabColor.Theme != null || sp.TabColor.Rgb != null)
                {
                    this.clrTabColor.FromTabColor(sp.TabColor);
                    HasTabColor = (clrTabColor.Color.IsEmpty) ? false : true;
                }
            }

            if (sp.OutlineProperties != null)
            {
                if (sp.OutlineProperties.ApplyStyles != null) this.ApplyStyles = sp.OutlineProperties.ApplyStyles.Value;
                if (sp.OutlineProperties.SummaryBelow != null) this.SummaryBelow = sp.OutlineProperties.SummaryBelow.Value;
                if (sp.OutlineProperties.SummaryRight != null) this.SummaryRight = sp.OutlineProperties.SummaryRight.Value;
                if (sp.OutlineProperties.ShowOutlineSymbols != null) this.ShowOutlineSymbols = sp.OutlineProperties.ShowOutlineSymbols.Value;
            }

            if (sp.PageSetupProperties != null)
            {
                if (sp.PageSetupProperties.AutoPageBreaks != null) this.AutoPageBreaks = sp.PageSetupProperties.AutoPageBreaks.Value;
                if (sp.PageSetupProperties.FitToPage != null) this.FitToPage = sp.PageSetupProperties.FitToPage.Value;
            }

            if (sp.SyncHorizontal != null) this.SyncHorizontal = sp.SyncHorizontal.Value;
            if (sp.SyncVertical != null) this.SyncVertical = sp.SyncVertical.Value;
            if (sp.SyncReference != null) this.SyncReference = sp.SyncReference.Value;
            if (sp.TransitionEvaluation != null) this.TransitionEvaluation = sp.TransitionEvaluation.Value;
            if (sp.TransitionEntry != null) this.TransitionEntry = sp.TransitionEntry.Value;
            if (sp.Published != null) this.Published = sp.Published.Value;
            if (sp.CodeName != null) this.CodeName = sp.CodeName.Value;
            if (sp.FilterMode != null) this.FilterMode = sp.FilterMode.Value;
            if (sp.EnableFormatConditionsCalculation != null) this.EnableFormatConditionsCalculation = sp.EnableFormatConditionsCalculation.Value;
        }

        internal SheetProperties ToSheetProperties()
        {
            SheetProperties sp = new SheetProperties();

            if (this.HasTabColor)
            {
                sp.TabColor = this.clrTabColor.ToTabColor();
            }

            if (this.ApplyStyles || !this.SummaryBelow || !this.SummaryRight || !this.ShowOutlineSymbols)
            {
                sp.OutlineProperties = new OutlineProperties();
                if (this.ApplyStyles) sp.OutlineProperties.ApplyStyles = this.ApplyStyles;
                if (!this.SummaryBelow) sp.OutlineProperties.SummaryBelow = this.SummaryBelow;
                if (!this.SummaryRight) sp.OutlineProperties.SummaryRight = this.SummaryRight;
                if (!this.ShowOutlineSymbols) sp.OutlineProperties.ShowOutlineSymbols = this.ShowOutlineSymbols;
            }

            if (!this.AutoPageBreaks || this.FitToPage)
            {
                sp.PageSetupProperties = new PageSetupProperties();
                if (!this.AutoPageBreaks) sp.PageSetupProperties.AutoPageBreaks = this.AutoPageBreaks;
                if (this.FitToPage) sp.PageSetupProperties.FitToPage = this.FitToPage;
            }

            if (this.SyncHorizontal) sp.SyncHorizontal = this.SyncHorizontal;
            if (this.SyncVertical) sp.SyncVertical = this.SyncVertical;
            if (this.SyncReference.Length > 0) sp.SyncReference = this.SyncReference;
            if (this.TransitionEvaluation) sp.TransitionEvaluation = this.TransitionEvaluation;
            if (this.TransitionEntry) sp.TransitionEntry = this.TransitionEntry;
            if (!this.Published) sp.Published = this.Published;
            if (this.CodeName.Length > 0) sp.CodeName = this.CodeName;
            if (this.FilterMode) sp.FilterMode = this.FilterMode;
            if (!this.EnableFormatConditionsCalculation) sp.EnableFormatConditionsCalculation = this.EnableFormatConditionsCalculation;

            return sp;
        }

        internal void FromChartSheetProperties(ChartSheetProperties sp)
        {
            this.SetAllNull();
            if (sp.TabColor != null)
            {
                if (sp.TabColor.Indexed != null || sp.TabColor.Theme != null || sp.TabColor.Rgb != null)
                {
                    this.clrTabColor.FromTabColor(sp.TabColor);
                    HasTabColor = (clrTabColor.Color.IsEmpty) ? false : true;
                }
            }

            if (sp.Published != null) this.Published = sp.Published.Value;
            if (sp.CodeName != null) this.CodeName = sp.CodeName.Value;
        }

        internal ChartSheetProperties ToChartSheetProperties()
        {
            ChartSheetProperties csp = new ChartSheetProperties();

            if (this.HasTabColor)
            {
                csp.TabColor = this.clrTabColor.ToTabColor();
            }

            if (!this.Published) csp.Published = this.Published;
            if (this.CodeName.Length > 0) csp.CodeName = this.CodeName;

            return csp;
        }

        internal SLSheetProperties Clone()
        {
            SLSheetProperties sp = new SLSheetProperties(this.listThemeColors, this.listIndexedColors);
            sp.clrTabColor = this.clrTabColor.Clone();
            sp.HasTabColor = this.HasTabColor;

            sp.ApplyStyles = this.ApplyStyles;
            sp.SummaryBelow = this.SummaryBelow;
            sp.SummaryRight = this.SummaryRight;
            sp.ShowOutlineSymbols = this.ShowOutlineSymbols;

            sp.AutoPageBreaks = this.AutoPageBreaks;
            sp.FitToPage = this.FitToPage;

            sp.SyncHorizontal = this.SyncHorizontal;
            sp.SyncVertical = this.SyncVertical;
            sp.SyncReference = this.SyncReference;
            sp.TransitionEvaluation = this.TransitionEvaluation;
            sp.TransitionEntry = this.TransitionEntry;
            sp.Published = this.Published;
            sp.CodeName = this.CodeName;
            sp.FilterMode = this.FilterMode;
            sp.EnableFormatConditionsCalculation = this.EnableFormatConditionsCalculation;

            return sp;
        }
    }
}
