using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLFieldGroup
    {
        internal bool HasRangeProperties;
        internal SLRangeProperties RangeProperties { get; set; }

        internal List<uint> DiscreteProperties { get; set; }

        internal bool HasGroupItems;
        internal SLGroupItems GroupItems { get; set; }

        internal uint? ParentId { get; set; }
        internal uint? Base { get; set; }

        internal SLFieldGroup()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.HasRangeProperties = false;
            this.RangeProperties = new SLRangeProperties();

            this.DiscreteProperties = new List<uint>();

            this.HasGroupItems = false;
            this.GroupItems = new SLGroupItems();

            this.ParentId = null;
            this.Base = null;
        }

        internal void FromFieldGroup(FieldGroup fg)
        {
            this.SetAllNull();

            if (fg.ParentId != null) this.ParentId = fg.ParentId.Value;
            if (fg.Base != null) this.Base = fg.Base.Value;

            using (OpenXmlReader oxr = OpenXmlReader.Create(fg))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(RangeProperties))
                    {
                        this.RangeProperties.FromRangeProperties((RangeProperties)oxr.LoadCurrentElement());
                        this.HasRangeProperties = true;
                    }
                    else if (oxr.ElementType == typeof(DiscreteProperties))
                    {
                        DiscreteProperties dp = (DiscreteProperties)oxr.LoadCurrentElement();
                        FieldItem fi;
                        using (OpenXmlReader oxrDiscrete = OpenXmlReader.Create(dp))
                        {
                            while (oxrDiscrete.Read())
                            {
                                if (oxrDiscrete.ElementType == typeof(FieldItem))
                                {
                                    fi = (FieldItem)oxrDiscrete.LoadCurrentElement();
                                    this.DiscreteProperties.Add(fi.Val);
                                }
                            }
                        }
                    }
                    else if (oxr.ElementType == typeof(GroupItems))
                    {
                        this.GroupItems.FromGroupItems((GroupItems)oxr.LoadCurrentElement());
                        this.HasGroupItems = true;
                    }
                }
            }
        }

        internal FieldGroup ToFieldGroup()
        {
            FieldGroup fg = new FieldGroup();
            if (this.ParentId != null) fg.ParentId = this.ParentId.Value;
            if (this.Base != null) fg.Base = this.Base.Value;

            if (this.HasRangeProperties)
            {
                fg.Append(this.RangeProperties.ToRangeProperties());
            }

            if (this.DiscreteProperties.Count > 0)
            {
                DiscreteProperties dp = new DiscreteProperties();
                dp.Count = (uint)this.DiscreteProperties.Count;
                foreach (uint i in this.DiscreteProperties)
                {
                    dp.Append(new FieldItem() { Val = i });
                }

                fg.Append(dp);
            }

            if (this.HasGroupItems)
            {
                fg.Append(this.GroupItems.ToGroupItems());
            }

            return fg;
        }

        internal SLFieldGroup Clone()
        {
            SLFieldGroup fg = new SLFieldGroup();
            fg.ParentId = this.ParentId;
            fg.Base = this.Base;

            fg.HasRangeProperties = this.HasRangeProperties;
            fg.RangeProperties = this.RangeProperties.Clone();

            fg.DiscreteProperties = new List<uint>();
            foreach (uint i in this.DiscreteProperties)
            {
                fg.DiscreteProperties.Add(i);
            }

            fg.HasGroupItems = this.HasGroupItems;
            fg.GroupItems = this.GroupItems.Clone();

            return fg;
        }
    }
}
