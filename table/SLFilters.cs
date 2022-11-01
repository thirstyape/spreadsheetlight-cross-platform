using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLFilters
    {
        internal List<SLFilter> Filters { get; set; }
        internal List<SLDateGroupItem> DateGroupItems { get; set; }
        internal bool? Blank { get; set; }

        internal bool HasCalendarType;
        private CalendarValues vCalendarType;
        internal CalendarValues CalendarType
        {
            get { return vCalendarType; }
            set
            {
                vCalendarType = value;
                HasCalendarType = vCalendarType != CalendarValues.None ? true : false;
            }
        }

        internal SLFilters()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Filters = new List<SLFilter>();
            this.DateGroupItems = new List<SLDateGroupItem>();
            this.Blank = null;
            this.vCalendarType = CalendarValues.None;
            this.HasCalendarType = false;
        }

        internal void FromFilters(Filters fs)
        {
            this.SetAllNull();

            if (fs.Blank != null && fs.Blank.Value) this.Blank = fs.Blank.Value;
            if (fs.CalendarType != null) this.CalendarType = fs.CalendarType.Value;

            if (fs.HasChildren)
            {
                SLFilter f;
                SLDateGroupItem dgi;
                using (OpenXmlReader oxr = OpenXmlReader.Create(fs))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(Filter))
                        {
                            f = new SLFilter();
                            f.FromFilter((Filter)oxr.LoadCurrentElement());
                            this.Filters.Add(f);
                        }
                        else if (oxr.ElementType == typeof(DateGroupItem))
                        {
                            dgi = new SLDateGroupItem();
                            dgi.FromDateGroupItem((DateGroupItem)oxr.LoadCurrentElement());
                            this.DateGroupItems.Add(dgi);
                        }
                    }
                }
            }
        }

        internal Filters ToFilters()
        {
            Filters fs = new Filters();
            if (this.Blank != null && this.Blank.Value) fs.Blank = this.Blank.Value;
            if (HasCalendarType) fs.CalendarType = this.CalendarType;

            foreach (SLFilter f in this.Filters)
            {
                fs.Append(f.ToFilter());
            }

            foreach (SLDateGroupItem dgi in this.DateGroupItems)
            {
                fs.Append(dgi.ToDateGroupItem());
            }

            return fs;
        }

        internal SLFilters Clone()
        {
            SLFilters fs = new SLFilters();

            int i;
            fs.Filters = new List<SLFilter>();
            for (i = 0; i < this.Filters.Count; ++i)
            {
                fs.Filters.Add(this.Filters[i].Clone());
            }

            fs.DateGroupItems = new List<SLDateGroupItem>();
            for (i = 0; i < this.DateGroupItems.Count; ++i)
            {
                fs.DateGroupItems.Add(this.DateGroupItems[i].Clone());
            }

            fs.Blank = this.Blank;
            fs.HasCalendarType = this.HasCalendarType;
            fs.vCalendarType = this.vCalendarType;

            return fs;
        }
    }
}
