using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLFilterColumn
    {
        internal bool HasFilters;
        internal SLFilters Filters { get; set; }

        internal bool HasTop10;
        internal SLTop10 Top10 { get; set; }

        internal bool HasCustomFilters;
        internal SLCustomFilters CustomFilters { get; set; }

        internal bool HasDynamicFilter;
        internal SLDynamicFilter DynamicFilter { get; set; }

        internal bool HasColorFilter;
        internal SLColorFilter ColorFilter { get; set; }

        internal bool HasIconFilter;
        internal SLIconFilter IconFilter { get; set; }

        internal uint ColumnId { get; set; }
        internal bool? HiddenButton { get; set; }
        internal bool? ShowButton { get; set; }

        internal SLFilterColumn()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Filters = new SLFilters();
            this.HasFilters = false;
            this.Top10 = new SLTop10();
            this.HasTop10 = false;
            this.CustomFilters = new SLCustomFilters();
            this.HasCustomFilters = false;
            this.DynamicFilter = new SLDynamicFilter();
            this.HasDynamicFilter = false;
            this.ColorFilter = new SLColorFilter();
            this.HasColorFilter = false;
            this.IconFilter = new SLIconFilter();
            this.HasIconFilter = false;
            this.ColumnId = 1;
            this.HiddenButton = null;
            this.ShowButton = null;
        }

        private void SetFiltersNull()
        {
            this.HasFilters = false;
            this.HasTop10 = false;
            this.HasCustomFilters = false;
            this.HasDynamicFilter = false;
            this.HasColorFilter = false;
            this.HasIconFilter = false;
        }

        internal void FromFilterColumn(FilterColumn fc)
        {
            this.SetAllNull();

            if (fc.Filters != null)
            {
                this.Filters.FromFilters(fc.Filters);
                this.HasFilters = true;
            }
            if (fc.Top10 != null)
            {
                this.Top10.FromTop10(fc.Top10);
                this.HasTop10 = true;
            }
            if (fc.CustomFilters != null)
            {
                this.CustomFilters.FromCustomFilters(fc.CustomFilters);
                this.HasCustomFilters = true;
            }
            if (fc.DynamicFilter != null)
            {
                this.DynamicFilter.FromDynamicFilter(fc.DynamicFilter);
                this.HasDynamicFilter = true;
            }
            if (fc.ColorFilter != null)
            {
                this.ColorFilter.FromColorFilter(fc.ColorFilter);
                this.HasColorFilter = true;
            }
            if (fc.IconFilter != null)
            {
                this.IconFilter.FromIconFilter(fc.IconFilter);
                this.HasIconFilter = true;
            }

            this.ColumnId = fc.ColumnId.Value;
            if (fc.HiddenButton != null && fc.HiddenButton.Value) this.HiddenButton = fc.HiddenButton.Value;
            if (fc.ShowButton != null && !fc.ShowButton.Value) this.ShowButton = fc.ShowButton.Value;
        }

        internal FilterColumn ToFilterColumn()
        {
            FilterColumn fc = new FilterColumn();

            if (HasFilters) fc.Filters = this.Filters.ToFilters();
            if (HasTop10) fc.Top10 = this.Top10.ToTop10();
            if (HasCustomFilters) fc.CustomFilters = this.CustomFilters.ToCustomFilters();
            if (HasDynamicFilter) fc.DynamicFilter = this.DynamicFilter.ToDynamicFilter();
            if (HasColorFilter) fc.ColorFilter = this.ColorFilter.ToColorFilter();
            if (HasIconFilter) fc.IconFilter = this.IconFilter.ToIconFilter();
            fc.ColumnId = this.ColumnId;
            if (this.HiddenButton != null && this.HiddenButton.Value) fc.HiddenButton = this.HiddenButton.Value;
            if (this.ShowButton != null && !this.ShowButton.Value) fc.ShowButton = this.ShowButton.Value;

            return fc;
        }

        internal SLFilterColumn Clone()
        {
            SLFilterColumn fc = new SLFilterColumn();
            fc.HasFilters = this.HasFilters;
            fc.Filters = this.Filters.Clone();
            fc.HasTop10 = this.HasTop10;
            fc.Top10 = this.Top10.Clone();
            fc.HasCustomFilters = this.HasCustomFilters;
            fc.CustomFilters = this.CustomFilters.Clone();
            fc.HasDynamicFilter = this.HasDynamicFilter;
            fc.DynamicFilter = this.DynamicFilter.Clone();
            fc.HasColorFilter = this.HasColorFilter;
            fc.ColorFilter = this.ColorFilter.Clone();
            fc.HasIconFilter = this.HasIconFilter;
            fc.IconFilter = this.IconFilter.Clone();
            fc.ColumnId = this.ColumnId;
            fc.HiddenButton = this.HiddenButton;
            fc.ShowButton = this.ShowButton;

            return fc;
        }
    }
}
