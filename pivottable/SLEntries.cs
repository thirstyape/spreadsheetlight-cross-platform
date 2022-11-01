using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal enum SLEntriesItemsType
    {
        Missing = 0,
        Number,
        Error,
        String
    }

    internal struct SLEntriesItemsTypeIndexPair
    {
        internal SLEntriesItemsType Type;
        // this is the 0-based index into whichever List<> depending on Type.
        internal int Index;

        internal SLEntriesItemsTypeIndexPair(SLEntriesItemsType Type, int Index)
        {
            this.Type = Type;
            this.Index = Index;
        }
    }

    internal class SLEntries
    {
        internal List<SLEntriesItemsTypeIndexPair> Items { get; set; }

        internal List<SLMissingItem> MissingItems { get; set; }
        internal List<SLNumberItem> NumberItems { get; set; }
        internal List<SLErrorItem> ErrorItems { get; set; }
        internal List<SLStringItem> StringItems { get; set; }

        internal SLEntries()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Items = new List<SLEntriesItemsTypeIndexPair>();

            this.MissingItems = new List<SLMissingItem>();
            this.NumberItems = new List<SLNumberItem>();
            this.ErrorItems = new List<SLErrorItem>();
            this.StringItems = new List<SLStringItem>();
        }

        internal void FromEntries(Entries es)
        {
            this.SetAllNull();

            SLMissingItem mi;
            SLNumberItem ni;
            SLErrorItem ei;
            SLStringItem si;
            using (OpenXmlReader oxr = OpenXmlReader.Create(es))
            {
                while (oxr.Read())
                {
                    // make sure to add to Items first, because of the Count thing.
                    if (oxr.ElementType == typeof(MissingItem))
                    {
                        mi = new SLMissingItem();
                        mi.FromMissingItem((MissingItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLEntriesItemsTypeIndexPair(SLEntriesItemsType.Missing, this.MissingItems.Count));
                        this.MissingItems.Add(mi);
                    }
                    else if (oxr.ElementType == typeof(NumberItem))
                    {
                        ni = new SLNumberItem();
                        ni.FromNumberItem((NumberItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLEntriesItemsTypeIndexPair(SLEntriesItemsType.Number, this.NumberItems.Count));
                        this.NumberItems.Add(ni);
                    }
                    else if (oxr.ElementType == typeof(ErrorItem))
                    {
                        ei = new SLErrorItem();
                        ei.FromErrorItem((ErrorItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLEntriesItemsTypeIndexPair(SLEntriesItemsType.Error, this.ErrorItems.Count));
                        this.ErrorItems.Add(ei);
                    }
                    else if (oxr.ElementType == typeof(StringItem))
                    {
                        si = new SLStringItem();
                        si.FromStringItem((StringItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLEntriesItemsTypeIndexPair(SLEntriesItemsType.String, this.StringItems.Count));
                        this.StringItems.Add(si);
                    }
                }
            }
        }

        internal Entries ToEntries()
        {
            Entries es = new Entries();

            // is it the sum of all the various items?
            if (this.Items.Count > 0) es.Count = (uint)this.Items.Count;

            foreach (SLEntriesItemsTypeIndexPair pair in this.Items)
            {
                switch (pair.Type)
                {
                    case SLEntriesItemsType.Missing:
                        es.Append(this.MissingItems[pair.Index].ToMissingItem());
                        break;
                    case SLEntriesItemsType.Number:
                        es.Append(this.NumberItems[pair.Index].ToNumberItem());
                        break;
                    case SLEntriesItemsType.Error:
                        es.Append(this.ErrorItems[pair.Index].ToErrorItem());
                        break;
                    case SLEntriesItemsType.String:
                        es.Append(this.StringItems[pair.Index].ToStringItem());
                        break;
                }
            }

            return es;
        }

        internal SLEntries Clone()
        {
            SLEntries es = new SLEntries();

            es.Items = new List<SLEntriesItemsTypeIndexPair>();
            foreach (SLEntriesItemsTypeIndexPair pair in this.Items)
            {
                es.Items.Add(new SLEntriesItemsTypeIndexPair(pair.Type, pair.Index));
            }

            es.MissingItems = new List<SLMissingItem>();
            foreach (SLMissingItem mi in this.MissingItems)
            {
                es.MissingItems.Add(mi.Clone());
            }

            es.NumberItems = new List<SLNumberItem>();
            foreach (SLNumberItem ni in this.NumberItems)
            {
                es.NumberItems.Add(ni.Clone());
            }

            es.ErrorItems = new List<SLErrorItem>();
            foreach (SLErrorItem ei in this.ErrorItems)
            {
                es.ErrorItems.Add(ei.Clone());
            }

            es.StringItems = new List<SLStringItem>();
            foreach (SLStringItem si in this.StringItems)
            {
                es.StringItems.Add(si.Clone());
            }

            return es;
        }
    }
}
