using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLConsolidation
    {
        internal List<List<string>> Pages { get; set; }
        internal List<SLRangeSet> RangeSets { get; set; }

        internal bool AutoPage { get; set; }

        internal SLConsolidation()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Pages = new List<List<string>>();
            this.RangeSets = new List<SLRangeSet>();
            this.AutoPage = true;
        }

        internal void FromConsolidation(Consolidation c)
        {
            this.SetAllNull();

            if (c.AutoPage != null) this.AutoPage = c.AutoPage.Value;

            Page pg;
            PageItem pgi;
            List<string> listPage;
            SLRangeSet rs;
            using (OpenXmlReader oxr = OpenXmlReader.Create(c))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Page))
                    {
                        listPage = new List<string>();
                        pg = (Page)oxr.LoadCurrentElement();
                        using (OpenXmlReader oxrPage = OpenXmlReader.Create(pg))
                        {
                            while (oxrPage.Read())
                            {
                                if (oxrPage.ElementType == typeof(PageItem))
                                {
                                    pgi = (PageItem)oxrPage.LoadCurrentElement();
                                    listPage.Add(pgi.Name.Value);
                                }
                            }
                        }
                        this.Pages.Add(listPage);
                    }
                    else if (oxr.ElementType == typeof(RangeSet))
                    {
                        rs = new SLRangeSet();
                        rs.FromRangeSet((RangeSet)oxr.LoadCurrentElement());
                        this.RangeSets.Add(rs);
                    }
                }
            }
        }

        internal Consolidation ToConsolidation()
        {
            Consolidation c = new Consolidation();
            if (this.AutoPage != true) c.AutoPage = this.AutoPage;

            if (this.Pages.Count > 0)
            {
                Page pg;
                c.Pages = new Pages() { Count = (uint)this.Pages.Count };
                foreach (List<string> ls in this.Pages)
                {
                    pg = new Page() { Count = (uint)ls.Count };
                    foreach (string s in ls)
                    {
                        pg.Append(new PageItem() { Name = s });
                    }
                    c.Pages.Append(pg);
                }
            }

            c.RangeSets = new RangeSets() { Count = (uint)this.RangeSets.Count };
            foreach (SLRangeSet rs in this.RangeSets)
            {
                c.RangeSets.Append(rs.ToRangeSet());
            }

            return c;
        }

        internal SLConsolidation Clone()
        {
            SLConsolidation c = new SLConsolidation();
            c.AutoPage = this.AutoPage;

            List<string> list;
            foreach (List<string> ls in this.Pages)
            {
                list = new List<string>();
                foreach (string s in ls)
                {
                    list.Add(s);
                }
                c.Pages.Add(list);
            }

            c.RangeSets = new List<SLRangeSet>();
            foreach (SLRangeSet rs in this.RangeSets)
            {
                c.RangeSets.Add(rs.Clone());
            }

            return c;
        }
    }
}
