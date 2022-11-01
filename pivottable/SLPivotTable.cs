using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Data field function values.
    /// </summary>
    public enum SLDataFieldFunctionValues
    {
        /// <summary>
        /// Average
        /// </summary>
        Average = 0,
        /// <summary>
        /// Count non-empty cells
        /// </summary>
        Count,
        /// <summary>
        /// Count numbers
        /// </summary>
        CountNumbers,
        /// <summary>
        /// Maximum
        /// </summary>
        Maximum,
        /// <summary>
        /// Minimum
        /// </summary>
        Minimum,
        /// <summary>
        /// Product
        /// </summary>
        Product,
        /// <summary>
        /// Standard deviation
        /// </summary>
        StandardDeviation,
        /// <summary>
        /// Sum
        /// </summary>
        Sum,
        /// <summary>
        /// Variance
        /// </summary>
        Variance
    }

    internal enum SLPivotFieldTypeValues
    {
        Filter = 0,
        Column,
        Row,
        Value,
        NotUsed
    }

    internal class SLPivotFieldType
    {
        // determines whether to use FieldIndex or FieldName
        internal bool IsNumericIndex { get; set; }
        internal int FieldIndex { get; set; }
        internal string FieldName { get; set; }
        internal SLPivotFieldTypeValues FieldType { get; set; }

        internal SLPivotFieldType()
        {
            this.IsNumericIndex = true;
            this.FieldIndex = 0;
            this.FieldName = string.Empty;
            this.FieldType = SLPivotFieldTypeValues.NotUsed;
        }
    }

    public class SLPivotTable
    {
        //CT_pivotTableDefinition
        //DocumentFormat.OpenXml.Spreadsheet.PivotTableDefinition

        //From Open XML specs: When encountering sheet boundaries, the PivotTable is truncated rather than wrapped, and as much as possible shall be shown.

        /*
         * <x:pivotCacheDefinition xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="Vincent" refreshedDate="41315.775251967592" createdVersion="5" refreshedVersion="5" minRefreshableVersion="3" recordCount="5" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:cacheSource type="worksheet">
    <x:worksheetSource ref="A1:D6" sheet="Sheet1" />
  </x:cacheSource>


<x:pivotCacheDefinition xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="Vincent" refreshedDate="41315.776955555557" createdVersion="5" refreshedVersion="5" minRefreshableVersion="3" recordCount="5" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:cacheSource type="worksheet">
    <x:worksheetSource name="Table1" />
  </x:cacheSource>
         * */

        // Pivot tables can have the same names as normal tables.

        internal bool IsValid
        {
            get
            {
                return DataRange.StartRowIndex >= 1 && DataRange.StartRowIndex <= SLConstants.RowLimit
                    && DataRange.StartColumnIndex >= 1 & DataRange.StartColumnIndex <= SLConstants.ColumnLimit
                    && DataRange.EndRowIndex >= 1 && DataRange.EndRowIndex <= SLConstants.RowLimit
                    && DataRange.EndColumnIndex >= 1 && DataRange.EndColumnIndex <= SLConstants.ColumnLimit;
            }
        }

        internal bool IsNewPivotTable;
        internal SLCellPointRange DataRange;

        /// <summary>
        /// If true, then SheetTableName is a table name. Otherwise it's a worksheet name.
        /// </summary>
        internal bool IsDataSourceTable;
        internal string SheetTableName;

        internal SLLocation Location { get; set; }
        internal List<SLPivotField> PivotFields { get; set; }
        internal List<int> RowFields { get; set; }
        internal List<SLRowItem> RowItems { get; set; }
        internal List<int> ColumnFields { get; set; }
        // ColumnItems use RowItem as children. Hey I'm not the one who designed this.
        internal List<SLRowItem> ColumnItems { get; set; }
        internal List<SLPageField> PageFields { get; set; }
        internal List<SLDataField> DataFields { get; set; }
        internal List<SLFormat> Formats { get; set; }
        internal List<SLConditionalFormat> ConditionalFormats { get; set; }
        internal List<SLChartFormat> ChartFormats { get; set; }
        internal List<SLPivotHierarchy> PivotHierarchies { get; set; }
        internal SLPivotTableStyle PivotTableStyle { get; set; }
        internal List<SLPivotFilter> PivotFilters { get; set; }
        internal List<int> RowHierarchiesUsage { get; set; }
        internal List<int> ColumnHierarchiesUsage { get; set; }

        // Oh my gamma rays the attributes are so *not* in accordance with the Open XML specs...
        //http://msdn.microsoft.com/en-us/library/ff532298%28v=office.12%29.aspx
        //http://msdn.microsoft.com/en-us/library/ff534910%28v=office.12%29.aspx

        //required attribute
        internal string Name { get; set; }
        //required attribute
        internal uint CacheId { get; set; }

        internal bool DataOnRows { get; set; }
        internal uint? DataPosition { get; set; }

        #region AG_AutoFormat
        internal uint? AutoFormatId { get; set; }
        internal bool? ApplyNumberFormats { get; set; }
        internal bool? ApplyBorderFormats { get; set; }
        internal bool? ApplyFontFormats { get; set; }
        internal bool? ApplyPatternFormats { get; set; }
        internal bool? ApplyAlignmentFormats { get; set; }
        internal bool? ApplyWidthHeightFormats { get; set; }
        #endregion

        //required attribute
        internal string DataCaption { get; set; }

        internal string GrandTotalCaption { get; set; }
        internal string ErrorCaption { get; set; }
        internal bool ShowError { get; set; }
        internal string MissingCaption { get; set; }
        internal bool ShowMissing { get; set; }
        internal string PageStyle { get; set; }
        internal string PivotTableStyleName { get; set; }
        internal string VacatedStyle { get; set; }
        internal string Tag { get; set; }
        internal byte UpdatedVersion { get; set; }
        internal byte MinRefreshableVersion { get; set; }
        internal bool AsteriskTotals { get; set; }
        internal bool ShowItems { get; set; }
        internal bool EditData { get; set; }
        internal bool DisableFieldList { get; set; }
        internal bool ShowCalculatedMembers { get; set; }
        internal bool VisualTotals { get; set; }
        internal bool ShowMultipleLabel { get; set; }
        internal bool ShowDataDropDown { get; set; }
        internal bool ShowDrill { get; set; }
        internal bool PrintDrill { get; set; }
        internal bool ShowMemberPropertyTips { get; set; }
        internal bool ShowDataTips { get; set; }
        internal bool EnableWizard { get; set; }
        internal bool EnableDrill { get; set; }
        internal bool EnableFieldProperties { get; set; }
        internal bool PreserveFormatting { get; set; }
        internal bool UseAutoFormatting { get; set; }
        internal uint PageWrap { get; set; }
        internal bool PageOverThenDown { get; set; }
        internal bool SubtotalHiddenItems { get; set; }
        internal bool RowGrandTotals { get; set; }
        internal bool ColumnGrandTotals { get; set; }
        internal bool FieldPrintTitles { get; set; }
        internal bool ItemPrintTitles { get; set; }
        internal bool MergeItem { get; set; }
        internal bool ShowDropZones { get; set; }
        internal byte CreatedVersion { get; set; }
        internal uint Indent { get; set; }
        internal bool ShowEmptyRow { get; set; }
        internal bool ShowEmptyColumn { get; set; }
        internal bool ShowHeaders { get; set; }
        internal bool Compact { get; set; }
        internal bool Outline { get; set; }
        internal bool OutlineData { get; set; }
        internal bool CompactData { get; set; }
        internal bool Published { get; set; }
        internal bool GridDropZones { get; set; }
        internal bool StopImmersiveUi { get; set; }
        internal bool MultipleFieldFilters { get; set; }
        internal uint ChartFormat { get; set; }
        internal string RowHeaderCaption { get; set; }
        internal string ColumnHeaderCaption { get; set; }
        internal bool FieldListSortAscending { get; set; }
        // what happened to mdxSubqueries? It's in the Open XML specs...
        // See http://msdn.microsoft.com/en-us/library/ff532298%28v=office.12%29.aspx
        // "Office does not use the mdxSubqueries attribute". Really? Hurrumph.
        internal bool CustomListSort { get; set; }

        // we store the instructions for which field to be for which type.
        // Then when we actually render the pivot table at the point of inserting
        // into the worksheet, we'll use this instruction set to get and juggle the
        // data.
        // We do it this way so that if any other worksheet operations are done,
        // say insert rows or setting different cell values that happen to be in the
        // jurisdiction of the pivot table range, we don't have to rejuggle our pivot
        // table data because we haven't done anything yet!
        internal List<SLPivotFieldType> FieldSettingInstructions { get; set; }

        internal SLPivotTable()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.IsNewPivotTable = true;
            this.DataRange = new SLCellPointRange(-1, -1, -1, -1);
            this.IsDataSourceTable = false;
            this.SheetTableName = string.Empty;

            // Excel 2013 uses 5 for attributes createdVersion and updatedVersion
            // and 3 for attribute minRefreshableVersion.
            // I don't know what earlier versions of Excel use because I've uninstalled
            // the earlier versions of Excel 2007 and 2010...
            // These three attributes are application dependent, which means technically they're
            // dependent on SpreadsheetLight.
            // For the sake of simplicity, I'm going to set all three attributes to the value of 3,
            // so Excel (whatever version) can handle it without being insufferable.

            this.Name = "";
            this.CacheId = 0;
            this.DataOnRows = false;
            this.DataPosition = null;

            this.AutoFormatId = null;
            this.ApplyNumberFormats = null;
            this.ApplyBorderFormats = null;
            this.ApplyFontFormats = null;
            this.ApplyPatternFormats = null;
            this.ApplyAlignmentFormats = null;
            this.ApplyWidthHeightFormats = null;

            this.DataCaption = "";
            this.GrandTotalCaption = "";
            this.ErrorCaption = "";
            this.ShowError = false;
            this.MissingCaption = "";
            this.ShowMissing = true;
            this.PageStyle = "";
            this.PivotTableStyleName = "";
            this.VacatedStyle = "";
            this.Tag = "";
            this.UpdatedVersion = 3; // supposed to default 0. See above.
            this.MinRefreshableVersion = 3; // supposed to default 0. See above.
            this.AsteriskTotals = false;
            this.ShowItems = true;
            this.EditData = false;
            this.DisableFieldList = false;
            this.ShowCalculatedMembers = true;
            this.VisualTotals = true;
            this.ShowMultipleLabel = true;
            this.ShowDataDropDown = true;
            this.ShowDrill = true;
            this.PrintDrill = false;
            this.ShowMemberPropertyTips = true;
            this.ShowDataTips = true;
            this.EnableWizard = true;
            this.EnableDrill = true;
            this.EnableFieldProperties = true;
            this.PreserveFormatting = true;
            this.UseAutoFormatting = false;
            this.PageWrap = 0;
            this.PageOverThenDown = false;
            this.SubtotalHiddenItems = false;
            this.RowGrandTotals = true;
            this.ColumnGrandTotals = true;
            this.FieldPrintTitles = false;
            this.ItemPrintTitles = false;
            this.MergeItem = false;
            this.ShowDropZones = true;
            this.CreatedVersion = 3; // supposed to default 0. See above.
            this.Indent = 1;
            this.ShowEmptyRow = false;
            this.ShowEmptyColumn = false;
            this.ShowHeaders = true;
            this.Compact = true;
            this.Outline = false;
            this.OutlineData = false;
            this.CompactData = true;
            this.Published = false;
            this.GridDropZones = false;
            this.StopImmersiveUi = true;
            this.MultipleFieldFilters = true;
            this.ChartFormat = 0;
            this.RowHeaderCaption = "";
            this.ColumnHeaderCaption = "";
            this.FieldListSortAscending = false;
            this.CustomListSort = true;
        }

        public void SetFilterField(int FieldIndex)
        {
            SLPivotFieldType pft = new SLPivotFieldType();
            pft.IsNumericIndex = true;
            pft.FieldIndex = FieldIndex;
            pft.FieldType = SLPivotFieldTypeValues.Filter;
            this.FieldSettingInstructions.Add(pft);
        }

        public void SetFilterField(string FieldName)
        {
            SLPivotFieldType pft = new SLPivotFieldType();
            pft.IsNumericIndex = false;
            pft.FieldName = FieldName;
            pft.FieldType = SLPivotFieldTypeValues.Filter;
            this.FieldSettingInstructions.Add(pft);
        }

        public void SetColumnField(int FieldIndex)
        {
            SLPivotFieldType pft = new SLPivotFieldType();
            pft.IsNumericIndex = true;
            pft.FieldIndex = FieldIndex;
            pft.FieldType = SLPivotFieldTypeValues.Column;
            this.FieldSettingInstructions.Add(pft);
        }

        public void SetColumnField(string FieldName)
        {
            SLPivotFieldType pft = new SLPivotFieldType();
            pft.IsNumericIndex = false;
            pft.FieldName = FieldName;
            pft.FieldType = SLPivotFieldTypeValues.Column;
            this.FieldSettingInstructions.Add(pft);
        }

        public void SetRowField(int FieldIndex)
        {
            SLPivotFieldType pft = new SLPivotFieldType();
            pft.IsNumericIndex = true;
            pft.FieldIndex = FieldIndex;
            pft.FieldType = SLPivotFieldTypeValues.Row;
            this.FieldSettingInstructions.Add(pft);
        }

        public void SetRowField(string FieldName)
        {
            SLPivotFieldType pft = new SLPivotFieldType();
            pft.IsNumericIndex = false;
            pft.FieldName = FieldName;
            pft.FieldType = SLPivotFieldTypeValues.Row;
            this.FieldSettingInstructions.Add(pft);
        }

        public void SetValueField(int FieldIndex)
        {
            SLPivotFieldType pft = new SLPivotFieldType();
            pft.IsNumericIndex = true;
            pft.FieldIndex = FieldIndex;
            pft.FieldType = SLPivotFieldTypeValues.Value;
            this.FieldSettingInstructions.Add(pft);
        }

        public void SetValueField(string FieldName)
        {
            SLPivotFieldType pft = new SLPivotFieldType();
            pft.IsNumericIndex = false;
            pft.FieldName = FieldName;
            pft.FieldType = SLPivotFieldTypeValues.Value;
            this.FieldSettingInstructions.Add(pft);
        }

        internal PivotTableDefinition ToPivotTableDefinition()
        {
            PivotTableDefinition ptd = new PivotTableDefinition();

            ptd.Name = this.Name;
            ptd.CacheId = this.CacheId;
            if (this.DataOnRows != false) ptd.DataOnRows = this.DataOnRows;
            if (this.DataPosition != null) ptd.DataPosition = this.DataPosition.Value;

            if (this.AutoFormatId != null) ptd.AutoFormatId = this.AutoFormatId.Value;
            if (this.ApplyNumberFormats != null) ptd.ApplyNumberFormats = this.ApplyNumberFormats.Value;
            if (this.ApplyBorderFormats != null) ptd.ApplyBorderFormats = this.ApplyBorderFormats.Value;
            if (this.ApplyFontFormats != null) ptd.ApplyFontFormats = this.ApplyFontFormats.Value;
            if (this.ApplyPatternFormats != null) ptd.ApplyPatternFormats = this.ApplyPatternFormats.Value;
            if (this.ApplyAlignmentFormats != null) ptd.ApplyAlignmentFormats = this.ApplyAlignmentFormats.Value;
            if (this.ApplyWidthHeightFormats != null) ptd.ApplyWidthHeightFormats = this.ApplyWidthHeightFormats.Value;

            if (this.DataCaption != null && this.DataCaption.Length > 0) ptd.DataCaption = this.DataCaption;
            if (this.GrandTotalCaption != null && this.GrandTotalCaption.Length > 0) ptd.GrandTotalCaption = this.GrandTotalCaption;
            if (this.ErrorCaption != null && this.ErrorCaption.Length > 0) ptd.ErrorCaption = this.ErrorCaption;
            if (this.ShowError != false) ptd.ShowError = this.ShowError;
            if (this.MissingCaption != null && this.MissingCaption.Length > 0) ptd.MissingCaption = this.MissingCaption;
            if (this.ShowMissing != true) ptd.ShowMissing = this.ShowMissing;
            if (this.PageStyle != null && this.PageStyle.Length > 0) ptd.PageStyle = this.PageStyle;
            if (this.PivotTableStyleName != null && this.PivotTableStyleName.Length > 0) ptd.PivotTableStyleName = this.PivotTableStyleName;
            if (this.VacatedStyle != null && this.VacatedStyle.Length > 0) ptd.VacatedStyle = this.VacatedStyle;
            if (this.Tag != null && this.Tag.Length > 0) ptd.Tag = this.Tag;
            if (this.UpdatedVersion != 0) ptd.UpdatedVersion = this.UpdatedVersion;
            if (this.MinRefreshableVersion != 0) ptd.MinRefreshableVersion = this.MinRefreshableVersion;
            if (this.AsteriskTotals != false) ptd.AsteriskTotals = this.AsteriskTotals;
            if (this.ShowItems != true) ptd.ShowItems = this.ShowItems;
            if (this.EditData != false) ptd.EditData = this.EditData;
            if (this.DisableFieldList != false) ptd.DisableFieldList = this.DisableFieldList;
            if (this.ShowCalculatedMembers != true) ptd.ShowCalculatedMembers = this.ShowCalculatedMembers;
            if (this.VisualTotals != true) ptd.VisualTotals = this.VisualTotals;
            if (this.ShowMultipleLabel != true) ptd.ShowMultipleLabel = this.ShowMultipleLabel;
            if (this.ShowDataDropDown != true) ptd.ShowDataDropDown = this.ShowDataDropDown;
            if (this.ShowDrill != true) ptd.ShowDrill = this.ShowDrill;
            if (this.PrintDrill != false) ptd.PrintDrill = this.PrintDrill;
            if (this.ShowMemberPropertyTips != true) ptd.ShowMemberPropertyTips = this.ShowMemberPropertyTips;
            if (this.ShowDataTips != true) ptd.ShowDataTips = this.ShowDataTips;
            if (this.EnableWizard != true) ptd.EnableWizard = this.EnableWizard;
            if (this.EnableDrill != true) ptd.EnableDrill = this.EnableDrill;
            if (this.EnableFieldProperties != true) ptd.EnableFieldProperties = this.EnableFieldProperties;
            if (this.PreserveFormatting != true) ptd.PreserveFormatting = this.PreserveFormatting;
            if (this.UseAutoFormatting != false) ptd.UseAutoFormatting = this.UseAutoFormatting;
            if (this.PageWrap != 0) ptd.PageWrap = this.PageWrap;
            if (this.PageOverThenDown != false) ptd.PageOverThenDown = this.PageOverThenDown;
            if (this.SubtotalHiddenItems != false) ptd.SubtotalHiddenItems = this.SubtotalHiddenItems;
            if (this.RowGrandTotals != true) ptd.RowGrandTotals = this.RowGrandTotals;
            if (this.ColumnGrandTotals != true) ptd.ColumnGrandTotals = this.ColumnGrandTotals;
            if (this.FieldPrintTitles != false) ptd.FieldPrintTitles = this.FieldPrintTitles;
            if (this.ItemPrintTitles != false) ptd.ItemPrintTitles = this.ItemPrintTitles;
            if (this.MergeItem != false) ptd.MergeItem = this.MergeItem;
            if (this.ShowDropZones != true) ptd.ShowDropZones = this.ShowDropZones;
            if (this.CreatedVersion != 0) ptd.CreatedVersion = this.CreatedVersion;
            if (this.Indent != 1) ptd.Indent = this.Indent;
            if (this.ShowEmptyRow != false) ptd.ShowEmptyRow = this.ShowEmptyRow;
            if (this.ShowEmptyColumn != false) ptd.ShowEmptyColumn = this.ShowEmptyColumn;
            if (this.ShowHeaders != true) ptd.ShowHeaders = this.ShowHeaders;
            if (this.Compact != true) ptd.Compact = this.Compact;
            if (this.Outline != false) ptd.Outline = this.Outline;
            if (this.OutlineData != false) ptd.OutlineData = this.OutlineData;
            if (this.CompactData != true) ptd.CompactData = this.CompactData;
            if (this.Published != false) ptd.Published = this.Published;
            if (this.GridDropZones != false) ptd.GridDropZones = this.GridDropZones;
            if (this.StopImmersiveUi != true) ptd.StopImmersiveUi = this.StopImmersiveUi;
            if (this.MultipleFieldFilters != true) ptd.MultipleFieldFilters = this.MultipleFieldFilters;
            if (this.ChartFormat != 0) ptd.ChartFormat = this.ChartFormat;
            if (this.RowHeaderCaption != null && this.RowHeaderCaption.Length > 0) ptd.RowHeaderCaption = this.RowHeaderCaption;
            if (this.ColumnHeaderCaption != null && this.ColumnHeaderCaption.Length > 0) ptd.ColumnHeaderCaption = this.ColumnHeaderCaption;
            if (this.FieldListSortAscending != false) ptd.FieldListSortAscending = this.FieldListSortAscending;
            if (this.CustomListSort != true) ptd.CustomListSort = this.CustomListSort;

            ptd.Location = this.Location.ToLocation();

            if (this.PivotFields.Count > 0)
            {
                ptd.PivotFields = new PivotFields() { Count = (uint)this.PivotFields.Count };
                foreach (SLPivotField pf in this.PivotFields)
                {
                    ptd.PivotFields.Append(pf.ToPivotField());
                }
            }

            if (this.RowFields.Count > 0)
            {
                ptd.RowFields = new RowFields() { Count = (uint)this.RowFields.Count };
                foreach (int i in this.RowFields)
                {
                    ptd.RowFields.Append(new Field() { Index = i });
                }
            }

            if (this.RowItems.Count > 0)
            {
                ptd.RowItems = new RowItems() { Count = (uint)this.RowItems.Count };
                foreach (SLRowItem ri in this.RowItems)
                {
                    ptd.RowItems.Append(ri.ToRowItem());
                }
            }

            if (this.ColumnFields.Count > 0)
            {
                ptd.ColumnFields = new ColumnFields() { Count = (uint)this.ColumnFields.Count };
                foreach (int i in this.ColumnFields)
                {
                    ptd.ColumnFields.Append(new Field() { Index = i });
                }
            }

            if (this.ColumnItems.Count > 0)
            {
                ptd.ColumnItems = new ColumnItems() { Count = (uint)this.ColumnItems.Count };
                foreach (SLRowItem ri in this.ColumnItems)
                {
                    ptd.ColumnItems.Append(ri.ToRowItem());
                }
            }

            if (this.PageFields.Count > 0)
            {
                ptd.PageFields = new PageFields() { Count = (uint)this.PageFields.Count };
                foreach (SLPageField pf in this.PageFields)
                {
                    ptd.PageFields.Append(pf.ToPageField());
                }
            }

            if (this.DataFields.Count > 0)
            {
                ptd.DataFields = new DataFields() { Count = (uint)this.DataFields.Count };
                foreach (SLDataField df in this.DataFields)
                {
                    ptd.DataFields.Append(df.ToDataField());
                }
            }

            if (this.Formats.Count > 0)
            {
                ptd.Formats = new Formats() { Count = (uint)this.Formats.Count };
                foreach (SLFormat f in this.Formats)
                {
                    ptd.Formats.Append(f.ToFormat());
                }
            }

            if (this.ConditionalFormats.Count > 0)
            {
                ptd.ConditionalFormats = new ConditionalFormats() { Count = (uint)this.ConditionalFormats.Count };
                foreach (SLConditionalFormat cf in this.ConditionalFormats)
                {
                    ptd.ConditionalFormats.Append(cf.ToConditionalFormat());
                }
            }

            if (this.ChartFormats.Count > 0)
            {
                ptd.ChartFormats = new ChartFormats() { Count = (uint)this.ChartFormats.Count };
                foreach (SLChartFormat cf in this.ChartFormats)
                {
                    ptd.ChartFormats.Append(cf.ToChartFormat());
                }
            }

            if (this.PivotHierarchies.Count > 0)
            {
                ptd.PivotHierarchies = new PivotHierarchies() { Count = (uint)this.PivotHierarchies.Count };
                foreach (SLPivotHierarchy ph in this.PivotHierarchies)
                {
                    ptd.PivotHierarchies.Append(ph.ToPivotHierarchy());
                }
            }

            ptd.PivotTableStyle = this.PivotTableStyle.ToPivotTableStyle();

            if (this.PivotFilters.Count > 0)
            {
                ptd.PivotFilters = new PivotFilters() { Count = (uint)this.PivotFilters.Count };
                foreach (SLPivotFilter pf in this.PivotFilters)
                {
                    ptd.PivotFilters.Append(pf.ToPivotFilter());
                }
            }

            if (this.RowHierarchiesUsage.Count > 0)
            {
                ptd.RowHierarchiesUsage = new RowHierarchiesUsage() { Count = (uint)this.RowHierarchiesUsage.Count };
                foreach (int i in this.RowHierarchiesUsage)
                {
                    ptd.RowHierarchiesUsage.Append(new RowHierarchyUsage() { Value = i });
                }
            }

            if (this.ColumnHierarchiesUsage.Count > 0)
            {
                ptd.ColumnHierarchiesUsage = new ColumnHierarchiesUsage() { Count = (uint)this.ColumnHierarchiesUsage.Count };
                foreach (int i in this.ColumnHierarchiesUsage)
                {
                    ptd.ColumnHierarchiesUsage.Append(new ColumnHierarchyUsage() { Value = i });
                }
            }

            return ptd;
        }
    }
}
