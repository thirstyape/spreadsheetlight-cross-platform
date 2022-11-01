using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal enum SLSharedGroupItemsType
    {
        Missing = 0,
        Number,
        Boolean,
        Error,
        String,
        DateTime
    }

    internal struct SLSharedGroupItemsTypeIndexPair
    {
        internal SLSharedGroupItemsType Type;
        // this is the 0-based index into whichever List<> depending on Type.
        internal int Index;

        internal SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType Type, int Index)
        {
            this.Type = Type;
            this.Index = Index;
        }
    }

    internal class SLSharedItems
    {
        internal List<SLSharedGroupItemsTypeIndexPair> Items { get; set; }

        internal List<SLMissingItem> MissingItems { get; set; }
        internal List<SLNumberItem> NumberItems { get; set; }
        internal List<SLBooleanItem> BooleanItems { get; set; }
        internal List<SLErrorItem> ErrorItems { get; set; }
        internal List<SLStringItem> StringItems { get; set; }
        internal List<SLDateTimeItem> DateTimeItems { get; set; }

        internal bool ContainsSemiMixedTypes { get; set; }
        internal bool ContainsNonDate { get; set; }
        internal bool ContainsDate { get; set; }
        internal bool ContainsString { get; set; }
        internal bool ContainsBlank { get; set; }
        internal bool ContainsMixedTypes { get; set; }
        internal bool ContainsNumber { get; set; }
        internal bool ContainsInteger { get; set; }
        internal double? MinValue { get; set; }
        internal double? MaxValue { get; set; }
        internal DateTime? MinDate { get; set; }
        internal DateTime? MaxDate { get; set; }
        //No need? internal uint? Count { get; set; }
        internal bool LongText { get; set; }

        internal SLSharedItems()
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

            this.ContainsSemiMixedTypes = true;
            this.ContainsNonDate = true;
            this.ContainsDate = false;
            this.ContainsString = true;
            this.ContainsBlank = false;
            this.ContainsMixedTypes = false;
            this.ContainsNumber = false;
            this.ContainsInteger = false;
            this.MinValue = null;
            this.MaxValue = null;
            this.MinDate = null;
            this.MaxDate = null;
            //this.Count = null;
            this.LongText = false;
        }

        internal void FromSharedItems(SharedItems sis)
        {
            this.SetAllNull();

            if (sis.ContainsSemiMixedTypes != null) this.ContainsSemiMixedTypes = sis.ContainsSemiMixedTypes.Value;
            if (sis.ContainsNonDate != null) this.ContainsNonDate = sis.ContainsNonDate.Value;
            if (sis.ContainsDate != null) this.ContainsDate = sis.ContainsDate.Value;
            if (sis.ContainsString != null) this.ContainsString = sis.ContainsString.Value;
            if (sis.ContainsBlank != null) this.ContainsBlank = sis.ContainsBlank.Value;
            if (sis.ContainsMixedTypes != null) this.ContainsMixedTypes = sis.ContainsMixedTypes.Value;
            if (sis.ContainsNumber != null) this.ContainsNumber = sis.ContainsNumber.Value;
            if (sis.ContainsInteger != null) this.ContainsInteger = sis.ContainsInteger.Value;
            if (sis.MinValue != null) this.MinValue = sis.MinValue.Value;
            if (sis.MaxValue != null) this.MaxValue = sis.MaxValue.Value;
            if (sis.MinDate != null) this.MinDate = sis.MinDate.Value;
            if (sis.MaxDate != null) this.MaxDate = sis.MaxDate.Value;
            //count
            if (sis.LongText != null) this.LongText = sis.LongText.Value;

            SLMissingItem mi;
            SLNumberItem ni;
            SLBooleanItem bi;
            SLErrorItem ei;
            SLStringItem si;
            SLDateTimeItem dti;
            using (OpenXmlReader oxr = OpenXmlReader.Create(sis))
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

        internal SharedItems ToSharedItems()
        {
            SharedItems sis = new SharedItems();
            if (this.ContainsSemiMixedTypes != true) sis.ContainsSemiMixedTypes = this.ContainsSemiMixedTypes;
            if (this.ContainsNonDate != true) sis.ContainsNonDate = this.ContainsNonDate;
            if (this.ContainsDate != false) sis.ContainsDate = this.ContainsDate;
            if (this.ContainsString != true) sis.ContainsString = this.ContainsString;
            if (this.ContainsBlank != false) sis.ContainsBlank = this.ContainsBlank;
            if (this.ContainsMixedTypes != false) sis.ContainsMixedTypes = this.ContainsMixedTypes;
            if (this.ContainsNumber != false) sis.ContainsNumber = this.ContainsNumber;
            if (this.ContainsInteger != false) sis.ContainsInteger = this.ContainsInteger;
            if (this.MinValue != null) sis.MinValue = this.MinValue.Value;
            if (this.MaxValue != null) sis.MaxValue = this.MaxValue.Value;
            if (this.MinDate != null) sis.MinDate = new DateTimeValue(this.MinDate.Value);
            if (this.MaxDate != null) sis.MaxDate = new DateTimeValue(this.MaxDate.Value);

            // is it the sum of all the various items?
            if (this.Items.Count > 0) sis.Count = (uint)this.Items.Count;

            if (this.LongText != false) sis.LongText = this.LongText;

            foreach (SLSharedGroupItemsTypeIndexPair pair in this.Items)
            {
                switch (pair.Type)
                {
                    case SLSharedGroupItemsType.Missing:
                        sis.Append(this.MissingItems[pair.Index].ToMissingItem());
                        break;
                    case SLSharedGroupItemsType.Number:
                        sis.Append(this.NumberItems[pair.Index].ToNumberItem());
                        break;
                    case SLSharedGroupItemsType.Boolean:
                        sis.Append(this.BooleanItems[pair.Index].ToBooleanItem());
                        break;
                    case SLSharedGroupItemsType.Error:
                        sis.Append(this.ErrorItems[pair.Index].ToErrorItem());
                        break;
                    case SLSharedGroupItemsType.String:
                        sis.Append(this.StringItems[pair.Index].ToStringItem());
                        break;
                    case SLSharedGroupItemsType.DateTime:
                        sis.Append(this.DateTimeItems[pair.Index].ToDateTimeItem());
                        break;
                }
            }

            return sis;
        }

        internal SLSharedItems Clone()
        {
            SLSharedItems sis = new SLSharedItems();
            sis.ContainsSemiMixedTypes = this.ContainsSemiMixedTypes;
            sis.ContainsNonDate = this.ContainsNonDate;
            sis.ContainsDate = this.ContainsDate;
            sis.ContainsString = this.ContainsString;
            sis.ContainsBlank = this.ContainsBlank;
            sis.ContainsMixedTypes = this.ContainsMixedTypes;
            sis.ContainsNumber = this.ContainsNumber;
            sis.ContainsInteger = this.ContainsInteger;
            sis.MinValue = this.MinValue;
            sis.MaxValue = this.MaxValue;
            sis.MinDate = this.MinDate;
            sis.MaxDate = this.MaxDate;
            //count
            sis.LongText = this.LongText;

            sis.Items = new List<SLSharedGroupItemsTypeIndexPair>();
            foreach (SLSharedGroupItemsTypeIndexPair pair in this.Items)
            {
                sis.Items.Add(new SLSharedGroupItemsTypeIndexPair(pair.Type, pair.Index));
            }

            sis.MissingItems = new List<SLMissingItem>();
            foreach (SLMissingItem mi in this.MissingItems)
            {
                sis.MissingItems.Add(mi.Clone());
            }

            sis.NumberItems = new List<SLNumberItem>();
            foreach (SLNumberItem ni in this.NumberItems)
            {
                sis.NumberItems.Add(ni.Clone());
            }

            sis.BooleanItems = new List<SLBooleanItem>();
            foreach (SLBooleanItem bi in this.BooleanItems)
            {
                sis.BooleanItems.Add(bi.Clone());
            }

            sis.ErrorItems = new List<SLErrorItem>();
            foreach (SLErrorItem ei in this.ErrorItems)
            {
                sis.ErrorItems.Add(ei.Clone());
            }

            sis.StringItems = new List<SLStringItem>();
            foreach (SLStringItem si in this.StringItems)
            {
                sis.StringItems.Add(si.Clone());
            }

            sis.DateTimeItems = new List<SLDateTimeItem>();
            foreach (SLDateTimeItem dti in this.DateTimeItems)
            {
                sis.DateTimeItems.Add(dti.Clone());
            }

            return sis;
        }
    }
}
