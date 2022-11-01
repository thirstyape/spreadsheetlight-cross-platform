using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLPivotField
    {
        internal List<SLItem> Items { get; set; }

        internal bool HasAutoSortScope;
        internal SLAutoSortScope AutoSortScope { get; set; }

        internal string Name { get; set; }
        internal PivotTableAxisValues? Axis { get; set; }
        internal bool DataField { get; set; }
        internal string SubtotalCaption { get; set; }
        internal bool ShowDropDowns { get; set; }
        internal bool HiddenLevel { get; set; }
        internal string UniqueMemberProperty { get; set; }
        internal bool Compact { get; set; }
        internal bool AllDrilled { get; set; }
        internal uint? NumberFormatId { get; set; }
        internal bool Outline { get; set; }
        internal bool SubtotalTop { get; set; }
        internal bool DragToRow { get; set; }
        internal bool DragToColumn { get; set; }
        internal bool MultipleItemSelectionAllowed { get; set; }
        internal bool DragToPage { get; set; }
        internal bool DragToData { get; set; }
        internal bool DragOff { get; set; }
        internal bool ShowAll { get; set; }
        internal bool InsertBlankRow { get; set; }
        internal bool ServerField { get; set; }
        internal bool InsertPageBreak { get; set; }
        internal bool AutoShow { get; set; }
        internal bool TopAutoShow { get; set; }
        internal bool HideNewItems { get; set; }
        internal bool MeasureFilter { get; set; }
        internal bool IncludeNewItemsInFilter { get; set; }
        internal uint ItemPageCount { get; set; }
        internal FieldSortValues SortType { get; set; }
        internal bool? DataSourceSort { get; set; }
        internal bool NonAutoSortDefault { get; set; }
        internal uint? RankBy { get; set; }
        internal bool DefaultSubtotal { get; set; }
        internal bool SumSubtotal { get; set; }
        internal bool CountASubtotal { get; set; }
        internal bool AverageSubTotal { get; set; }
        internal bool MaxSubtotal { get; set; }
        internal bool MinSubtotal { get; set; }
        internal bool ApplyProductInSubtotal { get; set; }
        internal bool CountSubtotal { get; set; }
        internal bool ApplyStandardDeviationInSubtotal { get; set; }
        internal bool ApplyStandardDeviationPInSubtotal { get; set; }
        internal bool ApplyVarianceInSubtotal { get; set; }
        internal bool ApplyVariancePInSubtotal { get; set; }
        internal bool ShowPropCell { get; set; }
        internal bool ShowPropertyTooltip { get; set; }
        internal bool ShowPropAsCaption { get; set; }
        internal bool DefaultAttributeDrillState { get; set; }

        internal SLPivotField()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Items = new List<SLItem>();

            this.AutoSortScope = new SLAutoSortScope();
            this.HasAutoSortScope = false;

            this.Name = "";
            this.Axis = null;
            this.DataField = false;
            this.SubtotalCaption = "";
            this.ShowDropDowns = true;
            this.HiddenLevel = false;
            this.UniqueMemberProperty = "";
            this.Compact = true;
            this.AllDrilled = false;
            this.NumberFormatId = null;
            this.Outline = true;
            this.SubtotalTop = true;
            this.DragToRow = true;
            this.DragToColumn = true;
            this.MultipleItemSelectionAllowed = false;
            this.DragToPage = true;
            this.DragToData = true;
            this.DragOff = true;
            this.ShowAll = true;
            this.InsertBlankRow = false;
            this.ServerField = false;
            this.InsertPageBreak = false;
            this.AutoShow = false;
            this.TopAutoShow = true;
            this.HideNewItems = false;
            this.MeasureFilter = false;
            this.IncludeNewItemsInFilter = false;
            this.ItemPageCount = 10;
            this.SortType = FieldSortValues.Manual;
            this.DataSourceSort = null;
            this.NonAutoSortDefault = false;
            this.RankBy = null;
            this.DefaultSubtotal = true;
            this.SumSubtotal = false;
            this.CountASubtotal = false;
            this.AverageSubTotal = false;
            this.MaxSubtotal = false;
            this.MinSubtotal = false;
            this.ApplyProductInSubtotal = false;
            this.CountSubtotal = false;
            this.ApplyStandardDeviationInSubtotal = false;
            this.ApplyStandardDeviationPInSubtotal = false;
            this.ApplyVarianceInSubtotal = false;
            this.ApplyVariancePInSubtotal = false;
            this.ShowPropCell = false;
            this.ShowPropertyTooltip = false;
            this.ShowPropAsCaption = false;
            this.DefaultAttributeDrillState = false;
        }

        internal void FromPivotField(PivotField pf)
        {
            this.SetAllNull();

            if (pf.Name != null) this.Name = pf.Name.Value;
            if (pf.Axis != null) this.Axis = pf.Axis.Value;
            if (pf.DataField != null) this.DataField = pf.DataField.Value;
            if (pf.SubtotalCaption != null) this.SubtotalCaption = pf.SubtotalCaption.Value;
            if (pf.ShowDropDowns != null) this.ShowDropDowns = pf.ShowDropDowns.Value;
            if (pf.HiddenLevel != null) this.HiddenLevel = pf.HiddenLevel.Value;
            if (pf.UniqueMemberProperty != null) this.UniqueMemberProperty = pf.UniqueMemberProperty.Value;
            if (pf.Compact != null) this.Compact = pf.Compact.Value;
            if (pf.AllDrilled != null) this.AllDrilled = pf.AllDrilled.Value;
            if (pf.NumberFormatId != null) this.NumberFormatId = pf.NumberFormatId.Value;
            if (pf.Outline != null) this.Outline = pf.Outline.Value;
            if (pf.SubtotalTop != null) this.SubtotalTop = pf.SubtotalTop.Value;
            if (pf.DragToRow != null) this.DragToRow = pf.DragToRow.Value;
            if (pf.DragToColumn != null) this.DragToColumn = pf.DragToColumn.Value;
            if (pf.MultipleItemSelectionAllowed != null) this.MultipleItemSelectionAllowed = pf.MultipleItemSelectionAllowed.Value;
            if (pf.DragToPage != null) this.DragToPage = pf.DragToPage.Value;
            if (pf.DragToData != null) this.DragToData = pf.DragToData.Value;
            if (pf.DragOff != null) this.DragOff = pf.DragOff.Value;
            if (pf.ShowAll != null) this.ShowAll = pf.ShowAll.Value;
            if (pf.InsertBlankRow != null) this.InsertBlankRow = pf.InsertBlankRow.Value;
            if (pf.ServerField != null) this.ServerField = pf.ServerField.Value;
            if (pf.InsertPageBreak != null) this.InsertPageBreak = pf.InsertPageBreak.Value;
            if (pf.AutoShow != null) this.AutoShow = pf.AutoShow.Value;
            if (pf.TopAutoShow != null) this.TopAutoShow = pf.TopAutoShow.Value;
            if (pf.HideNewItems != null) this.HideNewItems = pf.HideNewItems.Value;
            if (pf.MeasureFilter != null) this.MeasureFilter = pf.MeasureFilter.Value;
            if (pf.IncludeNewItemsInFilter != null) this.IncludeNewItemsInFilter = pf.IncludeNewItemsInFilter.Value;
            if (pf.ItemPageCount != null) this.ItemPageCount = pf.ItemPageCount.Value;
            if (pf.SortType != null) this.SortType = pf.SortType.Value;
            if (pf.DataSourceSort != null) this.DataSourceSort = pf.DataSourceSort.Value;
            if (pf.NonAutoSortDefault != null) this.NonAutoSortDefault = pf.NonAutoSortDefault.Value;
            if (pf.RankBy != null) this.RankBy = pf.RankBy.Value;
            if (pf.DefaultSubtotal != null) this.DefaultSubtotal = pf.DefaultSubtotal.Value;
            if (pf.SumSubtotal != null) this.SumSubtotal = pf.SumSubtotal.Value;
            if (pf.CountASubtotal != null) this.CountASubtotal = pf.CountASubtotal.Value;
            if (pf.AverageSubTotal != null) this.AverageSubTotal = pf.AverageSubTotal.Value;
            if (pf.MaxSubtotal != null) this.MaxSubtotal = pf.MaxSubtotal.Value;
            if (pf.MinSubtotal != null) this.MinSubtotal = pf.MinSubtotal.Value;
            if (pf.ApplyProductInSubtotal != null) this.ApplyProductInSubtotal = pf.ApplyProductInSubtotal.Value;
            if (pf.CountSubtotal != null) this.CountSubtotal = pf.CountSubtotal.Value;
            if (pf.ApplyStandardDeviationInSubtotal != null) this.ApplyStandardDeviationInSubtotal = pf.ApplyStandardDeviationInSubtotal.Value;
            if (pf.ApplyStandardDeviationPInSubtotal != null) this.ApplyStandardDeviationPInSubtotal = pf.ApplyStandardDeviationPInSubtotal.Value;
            if (pf.ApplyVarianceInSubtotal != null) this.ApplyVarianceInSubtotal = pf.ApplyVarianceInSubtotal.Value;
            if (pf.ApplyVariancePInSubtotal != null) this.ApplyVariancePInSubtotal = pf.ApplyVariancePInSubtotal.Value;
            if (pf.ShowPropCell != null) this.ShowPropCell = pf.ShowPropCell.Value;
            if (pf.ShowPropertyTooltip != null) this.ShowPropertyTooltip = pf.ShowPropertyTooltip.Value;
            if (pf.ShowPropAsCaption != null) this.ShowPropAsCaption = pf.ShowPropAsCaption.Value;
            if (pf.DefaultAttributeDrillState != null) this.DefaultAttributeDrillState = pf.DefaultAttributeDrillState.Value;

            SLItem it;
            using (OpenXmlReader oxr = OpenXmlReader.Create(pf))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Item))
                    {
                        it = new SLItem();
                        it.FromItem((Item)oxr.LoadCurrentElement());
                        this.Items.Add(it);
                    }
                    else if (oxr.ElementType == typeof(AutoSortScope))
                    {
                        this.AutoSortScope.FromAutoSortScope((AutoSortScope)oxr.LoadCurrentElement());
                        this.HasAutoSortScope = true;
                    }
                }
            }
        }

        internal PivotField ToPivotField()
        {
            PivotField pf = new PivotField();
            if (this.Name != null && this.Name.Length > 0) pf.Name = this.Name;
            if (this.Axis != null) pf.Axis = this.Axis.Value;
            if (this.DataField != false) pf.DataField = this.DataField;
            if (this.SubtotalCaption != null && this.SubtotalCaption.Length > 0) pf.SubtotalCaption = this.SubtotalCaption;
            if (this.ShowDropDowns != true) pf.ShowDropDowns = this.ShowDropDowns;
            if (this.HiddenLevel != false) pf.HiddenLevel = this.HiddenLevel;
            if (this.UniqueMemberProperty != null && this.UniqueMemberProperty.Length > 0) pf.UniqueMemberProperty = this.UniqueMemberProperty;
            if (this.Compact != true) pf.Compact = this.Compact;
            if (this.AllDrilled != false) pf.AllDrilled = this.AllDrilled;
            if (this.NumberFormatId != null) pf.NumberFormatId = this.NumberFormatId.Value;
            if (this.Outline != true) pf.Outline = this.Outline;
            if (this.SubtotalTop != true) pf.SubtotalTop = this.SubtotalTop;
            if (this.DragToRow != true) pf.DragToRow = this.DragToRow;
            if (this.DragToColumn != true) pf.DragToColumn = this.DragToColumn;
            if (this.MultipleItemSelectionAllowed != false) pf.MultipleItemSelectionAllowed = this.MultipleItemSelectionAllowed;
            if (this.DragToPage != true) pf.DragToPage = this.DragToPage;
            if (this.DragToData != true) pf.DragToData = this.DragToData;
            if (this.DragOff != true) pf.DragOff = this.DragOff;
            if (this.ShowAll != true) pf.ShowAll = this.ShowAll;
            if (this.InsertBlankRow != false) pf.InsertBlankRow = this.InsertBlankRow;
            if (this.ServerField != false) pf.ServerField = this.ServerField;
            if (this.InsertPageBreak != false) pf.InsertPageBreak = this.InsertPageBreak;
            if (this.AutoShow != false) pf.AutoShow = this.AutoShow;
            if (this.TopAutoShow != true) pf.TopAutoShow = this.TopAutoShow;
            if (this.HideNewItems != false) pf.HideNewItems = this.HideNewItems;
            if (this.MeasureFilter != false) pf.MeasureFilter = this.MeasureFilter;
            if (this.IncludeNewItemsInFilter != false) pf.IncludeNewItemsInFilter = this.IncludeNewItemsInFilter;
            if (this.ItemPageCount != 10) pf.ItemPageCount = this.ItemPageCount;
            if (this.SortType != FieldSortValues.Manual) pf.SortType = this.SortType;
            if (this.DataSourceSort != null) pf.DataSourceSort = this.DataSourceSort.Value;
            if (this.NonAutoSortDefault != false) pf.NonAutoSortDefault = this.NonAutoSortDefault;
            if (this.RankBy != null) pf.RankBy = this.RankBy.Value;
            if (this.DefaultSubtotal != true) pf.DefaultSubtotal = this.DefaultSubtotal;
            if (this.SumSubtotal != false) pf.SumSubtotal = this.SumSubtotal;
            if (this.CountASubtotal != false) pf.CountASubtotal = this.CountASubtotal;
            if (this.AverageSubTotal != false) pf.AverageSubTotal = this.AverageSubTotal;
            if (this.MaxSubtotal != false) pf.MaxSubtotal = this.MaxSubtotal;
            if (this.MinSubtotal != false) pf.MinSubtotal = this.MinSubtotal;
            if (this.ApplyProductInSubtotal != false) pf.ApplyProductInSubtotal = this.ApplyProductInSubtotal;
            if (this.CountSubtotal != false) pf.CountSubtotal = this.CountSubtotal;
            if (this.ApplyStandardDeviationInSubtotal != false) pf.ApplyStandardDeviationInSubtotal = this.ApplyStandardDeviationInSubtotal;
            if (this.ApplyStandardDeviationPInSubtotal != false) pf.ApplyStandardDeviationPInSubtotal = this.ApplyStandardDeviationPInSubtotal;
            if (this.ApplyVarianceInSubtotal != false) pf.ApplyVarianceInSubtotal = this.ApplyVarianceInSubtotal;
            if (this.ApplyVariancePInSubtotal != false) pf.ApplyVariancePInSubtotal = this.ApplyVariancePInSubtotal;
            if (this.ShowPropCell != false) pf.ShowPropCell = this.ShowPropCell;
            if (this.ShowPropertyTooltip != false) pf.ShowPropertyTooltip = this.ShowPropertyTooltip;
            if (this.ShowPropAsCaption != false) pf.ShowPropAsCaption = this.ShowPropAsCaption;
            if (this.DefaultAttributeDrillState != false) pf.DefaultAttributeDrillState = this.DefaultAttributeDrillState;

            if (this.Items.Count > 0)
            {
                pf.Items = new Items();
                foreach (SLItem it in this.Items)
                {
                    pf.Items.Append(it.ToItem());
                }
            }

            if (this.HasAutoSortScope)
            {
                pf.AutoSortScope = this.AutoSortScope.ToAutoSortScope();
            }

            return pf;
        }

        internal SLPivotField Clone()
        {
            SLPivotField pf = new SLPivotField();
            pf.Name = this.Name;
            pf.Axis = this.Axis;
            pf.DataField = this.DataField;
            pf.SubtotalCaption = this.SubtotalCaption;
            pf.ShowDropDowns = this.ShowDropDowns;
            pf.HiddenLevel = this.HiddenLevel;
            pf.UniqueMemberProperty = this.UniqueMemberProperty;
            pf.Compact = this.Compact;
            pf.AllDrilled = this.AllDrilled;
            pf.NumberFormatId = this.NumberFormatId;
            pf.Outline = this.Outline;
            pf.SubtotalTop = this.SubtotalTop;
            pf.DragToRow = this.DragToRow;
            pf.DragToColumn = this.DragToColumn;
            pf.MultipleItemSelectionAllowed = this.MultipleItemSelectionAllowed;
            pf.DragToPage = this.DragToPage;
            pf.DragToData = this.DragToData;
            pf.DragOff = this.DragOff;
            pf.ShowAll = this.ShowAll;
            pf.InsertBlankRow = this.InsertBlankRow;
            pf.ServerField = this.ServerField;
            pf.InsertPageBreak = this.InsertPageBreak;
            pf.AutoShow = this.AutoShow;
            pf.TopAutoShow = this.TopAutoShow;
            pf.HideNewItems = this.HideNewItems;
            pf.MeasureFilter = this.MeasureFilter;
            pf.IncludeNewItemsInFilter = this.IncludeNewItemsInFilter;
            pf.ItemPageCount = this.ItemPageCount;
            pf.SortType = this.SortType;
            pf.DataSourceSort = this.DataSourceSort;
            pf.NonAutoSortDefault = this.NonAutoSortDefault;
            pf.RankBy = this.RankBy;
            pf.DefaultSubtotal = this.DefaultSubtotal;
            pf.SumSubtotal = this.SumSubtotal;
            pf.CountASubtotal = this.CountASubtotal;
            pf.AverageSubTotal = this.AverageSubTotal;
            pf.MaxSubtotal = this.MaxSubtotal;
            pf.MinSubtotal = this.MinSubtotal;
            pf.ApplyProductInSubtotal = this.ApplyProductInSubtotal;
            pf.CountSubtotal = this.CountSubtotal;
            pf.ApplyStandardDeviationInSubtotal = this.ApplyStandardDeviationInSubtotal;
            pf.ApplyStandardDeviationPInSubtotal = this.ApplyStandardDeviationPInSubtotal;
            pf.ApplyVarianceInSubtotal = this.ApplyVarianceInSubtotal;
            pf.ApplyVariancePInSubtotal = this.ApplyVariancePInSubtotal;
            pf.ShowPropCell = this.ShowPropCell;
            pf.ShowPropertyTooltip = this.ShowPropertyTooltip;
            pf.ShowPropAsCaption = this.ShowPropAsCaption;
            pf.DefaultAttributeDrillState = this.DefaultAttributeDrillState;

            pf.Items = new List<SLItem>();
            foreach (SLItem it in this.Items)
            {
                pf.Items.Add(it.Clone());
            }

            pf.AutoSortScope = this.AutoSortScope.Clone();
            pf.HasAutoSortScope = this.HasAutoSortScope;

            return pf;
        }
    }
}
