using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLGroupItems
    {
        internal List<SLSharedGroupItemsTypeIndexPair> Items { get; set; }

        internal List<SLMissingItem> MissingItems { get; set; }
        internal List<SLNumberItem> NumberItems { get; set; }
        internal List<SLBooleanItem> BooleanItems { get; set; }
        internal List<SLErrorItem> ErrorItems { get; set; }
        internal List<SLStringItem> StringItems { get; set; }
        internal List<SLDateTimeItem> DateTimeItems { get; set; }

        internal SLGroupItems()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Items = new List<SLSharedGroupItemsTypeIndexPair>();

            this.MissingItems = new List<SLMissingItem>();
            this.NumberItems = new List<SLNumberItem>();
            this.BooleanItems = new List<SLBooleanItem>();
            this.ErrorItems = new List<SLErrorItem>();
            this.StringItems = new List<SLStringItem>();
            this.DateTimeItems = new List<SLDateTimeItem>();
        }

        internal void FromGroupItems(GroupItems gis)
        {
            this.SetAllNull();

            SLMissingItem mi;
            SLNumberItem ni;
            SLBooleanItem bi;
            SLErrorItem ei;
            SLStringItem si;
            SLDateTimeItem dti;
            using (OpenXmlReader oxr = OpenXmlReader.Create(gis))
            {
                while (oxr.Read())
                {
                    // make sure to add to Items first, because of the Count thing.
                    if (oxr.ElementType == typeof(MissingItem))
                    {
                        mi = new SLMissingItem();
                        mi.FromMissingItem((MissingItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.Missing, this.MissingItems.Count));
                        this.MissingItems.Add(mi);
                    }
                    else if (oxr.ElementType == typeof(NumberItem))
                    {
                        ni = new SLNumberItem();
                        ni.FromNumberItem((NumberItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.Number, this.NumberItems.Count));
                        this.NumberItems.Add(ni);
                    }
                    else if (oxr.ElementType == typeof(BooleanItem))
                    {
                        bi = new SLBooleanItem();
                        bi.FromBooleanItem((BooleanItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.Boolean, this.BooleanItems.Count));
                        this.BooleanItems.Add(bi);
                    }
                    else if (oxr.ElementType == typeof(ErrorItem))
                    {
                        ei = new SLErrorItem();
                        ei.FromErrorItem((ErrorItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.Error, this.ErrorItems.Count));
                        this.ErrorItems.Add(ei);
                    }
                    else if (oxr.ElementType == typeof(StringItem))
                    {
                        si = new SLStringItem();
                        si.FromStringItem((StringItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.String, this.StringItems.Count));
                        this.StringItems.Add(si);
                    }
                    else if (oxr.ElementType == typeof(DateTimeItem))
                    {
                        dti = new SLDateTimeItem();
                        dti.FromDateTimeItem((DateTimeItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.DateTime, this.DateTimeItems.Count));
                        this.DateTimeItems.Add(dti);
                    }
                }
            }
        }

        internal GroupItems ToGroupItems()
        {
            GroupItems gis = new GroupItems();
            gis.Count = (uint)this.Items.Count;

            foreach (SLSharedGroupItemsTypeIndexPair pair in this.Items)
            {
                switch (pair.Type)
                {
                    case SLSharedGroupItemsType.Missing:
                        gis.Append(this.MissingItems[pair.Index].ToMissingItem());
                        break;
                    case SLSharedGroupItemsType.Number:
                        gis.Append(this.NumberItems[pair.Index].ToNumberItem());
                        break;
                    case SLSharedGroupItemsType.Boolean:
                        gis.Append(this.BooleanItems[pair.Index].ToBooleanItem());
                        break;
                    case SLSharedGroupItemsType.Error:
                        gis.Append(this.ErrorItems[pair.Index].ToErrorItem());
                        break;
                    case SLSharedGroupItemsType.String:
                        gis.Append(this.StringItems[pair.Index].ToStringItem());
                        break;
                    case SLSharedGroupItemsType.DateTime:
                        gis.Append(this.DateTimeItems[pair.Index].ToDateTimeItem());
                        break;
                }
            }

            return gis;
        }

        internal SLGroupItems Clone()
        {
            SLGroupItems gis = new SLGroupItems();

            gis.Items = new List<SLSharedGroupItemsTypeIndexPair>();
            foreach (SLSharedGroupItemsTypeIndexPair pair in this.Items)
            {
                gis.Items.Add(new SLSharedGroupItemsTypeIndexPair(pair.Type, pair.Index));
            }

            gis.MissingItems = new List<SLMissingItem>();
            foreach (SLMissingItem mi in this.MissingItems)
            {
                gis.MissingItems.Add(mi.Clone());
            }

            gis.NumberItems = new List<SLNumberItem>();
            foreach (SLNumberItem ni in this.NumberItems)
            {
                gis.NumberItems.Add(ni.Clone());
            }

            gis.BooleanItems = new List<SLBooleanItem>();
            foreach (SLBooleanItem bi in this.BooleanItems)
            {
                gis.BooleanItems.Add(bi.Clone());
            }

            gis.ErrorItems = new List<SLErrorItem>();
            foreach (SLErrorItem ei in this.ErrorItems)
            {
                gis.ErrorItems.Add(ei.Clone());
            }

            gis.StringItems = new List<SLStringItem>();
            foreach (SLStringItem si in this.StringItems)
            {
                gis.StringItems.Add(si.Clone());
            }

            gis.DateTimeItems = new List<SLDateTimeItem>();
            foreach (SLDateTimeItem dti in this.DateTimeItems)
            {
                gis.DateTimeItems.Add(dti.Clone());
            }

            return gis;
        }
    }
}
