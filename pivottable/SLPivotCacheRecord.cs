using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal enum SLPivotCacheRecordItemsType
    {
        Missing = 0,
        Number,
        Boolean,
        Error,
        String,
        DateTime,
        Field
    }

    internal struct SLPivotCacheRecordItemsTypeIndexPair
    {
        internal SLPivotCacheRecordItemsType Type;
        // this is the 0-based index into whichever List<> depending on Type.
        internal int Index;

        internal SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType Type, int Index)
        {
            this.Type = Type;
            this.Index = Index;
        }
    }

    internal class SLPivotCacheRecord
    {
        internal List<SLPivotCacheRecordItemsTypeIndexPair> Items { get; set; }

        internal List<SLMissingItem> MissingItems { get; set; }
        internal List<SLNumberItem> NumberItems { get; set; }
        internal List<SLBooleanItem> BooleanItems { get; set; }
        internal List<SLErrorItem> ErrorItems { get; set; }
        internal List<SLStringItem> StringItems { get; set; }
        internal List<SLDateTimeItem> DateTimeItems { get; set; }
        internal List<uint> FieldItems { get; set; }

        internal SLPivotCacheRecord()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Items = new List<SLPivotCacheRecordItemsTypeIndexPair>();

            this.MissingItems = new List<SLMissingItem>();
            this.NumberItems = new List<SLNumberItem>();
            this.BooleanItems = new List<SLBooleanItem>();
            this.ErrorItems = new List<SLErrorItem>();
            this.StringItems = new List<SLStringItem>();
            this.DateTimeItems = new List<SLDateTimeItem>();
            this.FieldItems = new List<uint>();
        }

        internal void FromPivotCacheRecord(PivotCacheRecord pcr)
        {
            this.SetAllNull();

            SLMissingItem mi;
            SLNumberItem ni;
            SLBooleanItem bi;
            SLErrorItem ei;
            SLStringItem si;
            SLDateTimeItem dti;
            FieldItem fi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(pcr))
            {
                while (oxr.Read())
                {
                    // make sure to add to Items first, because of the Count thing.
                    if (oxr.ElementType == typeof(MissingItem))
                    {
                        mi = new SLMissingItem();
                        mi.FromMissingItem((MissingItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Missing, this.MissingItems.Count));
                        this.MissingItems.Add(mi);
                    }
                    else if (oxr.ElementType == typeof(NumberItem))
                    {
                        ni = new SLNumberItem();
                        ni.FromNumberItem((NumberItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Number, this.NumberItems.Count));
                        this.NumberItems.Add(ni);
                    }
                    else if (oxr.ElementType == typeof(BooleanItem))
                    {
                        bi = new SLBooleanItem();
                        bi.FromBooleanItem((BooleanItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Boolean, this.BooleanItems.Count));
                        this.BooleanItems.Add(bi);
                    }
                    else if (oxr.ElementType == typeof(ErrorItem))
                    {
                        ei = new SLErrorItem();
                        ei.FromErrorItem((ErrorItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Error, this.ErrorItems.Count));
                        this.ErrorItems.Add(ei);
                    }
                    else if (oxr.ElementType == typeof(StringItem))
                    {
                        si = new SLStringItem();
                        si.FromStringItem((StringItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.String, this.StringItems.Count));
                        this.StringItems.Add(si);
                    }
                    else if (oxr.ElementType == typeof(DateTimeItem))
                    {
                        dti = new SLDateTimeItem();
                        dti.FromDateTimeItem((DateTimeItem)oxr.LoadCurrentElement());
                        this.Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.DateTime, this.DateTimeItems.Count));
                        this.DateTimeItems.Add(dti);
                    }
                    else if (oxr.ElementType == typeof(FieldItem))
                    {
                        fi = (FieldItem)oxr.LoadCurrentElement();
                        this.Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Field, this.FieldItems.Count));
                        this.FieldItems.Add(fi.Val.Value);
                    }
                }
            }
        }

        internal PivotCacheRecord ToPivotCacheRecord()
        {
            PivotCacheRecord pcr = new PivotCacheRecord();

            foreach (SLPivotCacheRecordItemsTypeIndexPair pair in this.Items)
            {
                switch (pair.Type)
                {
                    case SLPivotCacheRecordItemsType.Missing:
                        pcr.Append(this.MissingItems[pair.Index].ToMissingItem());
                        break;
                    case SLPivotCacheRecordItemsType.Number:
                        pcr.Append(this.NumberItems[pair.Index].ToNumberItem());
                        break;
                    case SLPivotCacheRecordItemsType.Boolean:
                        pcr.Append(this.BooleanItems[pair.Index].ToBooleanItem());
                        break;
                    case SLPivotCacheRecordItemsType.Error:
                        pcr.Append(this.ErrorItems[pair.Index].ToErrorItem());
                        break;
                    case SLPivotCacheRecordItemsType.String:
                        pcr.Append(this.StringItems[pair.Index].ToStringItem());
                        break;
                    case SLPivotCacheRecordItemsType.DateTime:
                        pcr.Append(this.DateTimeItems[pair.Index].ToDateTimeItem());
                        break;
                    case SLPivotCacheRecordItemsType.Field:
                        pcr.Append(new FieldItem() { Val = this.FieldItems[pair.Index] });
                        break;
                }
            }

            return pcr;
        }

        internal SLPivotCacheRecord Clone()
        {
            SLPivotCacheRecord pcr = new SLPivotCacheRecord();

            pcr.Items = new List<SLPivotCacheRecordItemsTypeIndexPair>();
            foreach (SLPivotCacheRecordItemsTypeIndexPair pair in this.Items)
            {
                pcr.Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(pair.Type, pair.Index));
            }

            pcr.MissingItems = new List<SLMissingItem>();
            foreach (SLMissingItem mi in this.MissingItems)
            {
                pcr.MissingItems.Add(mi.Clone());
            }

            pcr.NumberItems = new List<SLNumberItem>();
            foreach (SLNumberItem ni in this.NumberItems)
            {
                pcr.NumberItems.Add(ni.Clone());
            }

            pcr.BooleanItems = new List<SLBooleanItem>();
            foreach (SLBooleanItem bi in this.BooleanItems)
            {
                pcr.BooleanItems.Add(bi.Clone());
            }

            pcr.ErrorItems = new List<SLErrorItem>();
            foreach (SLErrorItem ei in this.ErrorItems)
            {
                pcr.ErrorItems.Add(ei.Clone());
            }

            pcr.StringItems = new List<SLStringItem>();
            foreach (SLStringItem si in this.StringItems)
            {
                pcr.StringItems.Add(si.Clone());
            }

            pcr.DateTimeItems = new List<SLDateTimeItem>();
            foreach (SLDateTimeItem dti in this.DateTimeItems)
            {
                pcr.DateTimeItems.Add(dti.Clone());
            }

            pcr.FieldItems = new List<uint>();
            foreach (uint i in this.FieldItems)
            {
                pcr.FieldItems.Add(i);
            }

            return pcr;
        }
    }
}
