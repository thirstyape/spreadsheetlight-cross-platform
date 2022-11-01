using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLCacheField
    {
        internal bool HasSharedItems;
        internal SLSharedItems SharedItems { get; set; }

        internal bool HasFieldGroup;
        internal SLFieldGroup FieldGroup { get; set; }

        internal List<int> MemberPropertiesMaps { get; set; }

        internal string Name { get; set; }
        internal string Caption { get; set; }
        internal string PropertyName { get; set; }
        internal bool ServerField { get; set; }
        internal bool UniqueList { get; set; }
        internal uint? NumberFormatId { get; set; }
        internal string Formula { get; set; }
        internal int SqlType { get; set; }
        internal int Hierarchy { get; set; }
        internal uint Level { get; set; }
        internal bool DatabaseField { get; set; }
        internal uint? MappingCount { get; set; }
        internal bool MemberPropertyField { get; set; }

        internal SLCacheField()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.HasSharedItems = false;
            this.SharedItems = new SLSharedItems();

            this.HasFieldGroup = false;
            this.FieldGroup = new SLFieldGroup();

            this.MemberPropertiesMaps = new List<int>();

            this.Name = "";
            this.Caption = "";
            this.PropertyName = "";
            this.ServerField = false;
            this.UniqueList = true;
            this.NumberFormatId = null;
            this.Formula = "";
            this.SqlType = 0;
            this.Hierarchy = 0;
            this.Level = 0;
            this.DatabaseField = true;
            this.MappingCount = null;
            this.MemberPropertyField = false;
        }

        internal void FromCacheField(CacheField cf)
        {
            this.SetAllNull();

            if (cf.Name != null) this.Name = cf.Name.Value;
            if (cf.Caption != null) this.Caption = cf.Caption.Value;
            if (cf.PropertyName != null) this.PropertyName = cf.PropertyName.Value;
            if (cf.ServerField != null) this.ServerField = cf.ServerField.Value;
            if (cf.UniqueList != null) this.UniqueList = cf.UniqueList.Value;
            if (cf.NumberFormatId != null) this.NumberFormatId = cf.NumberFormatId.Value;
            if (cf.Formula != null) this.Formula = cf.Formula.Value;
            if (cf.SqlType != null) this.SqlType = cf.SqlType.Value;
            if (cf.Hierarchy != null) this.Hierarchy = cf.Hierarchy.Value;
            if (cf.Level != null) this.Level = cf.Level.Value;
            if (cf.DatabaseField != null) this.DatabaseField = cf.DatabaseField.Value;
            if (cf.MappingCount != null) this.MappingCount = cf.MappingCount.Value;
            if (cf.MemberPropertyField != null) this.MemberPropertyField = cf.MemberPropertyField.Value;

            if (cf.SharedItems != null)
            {
                this.SharedItems.FromSharedItems(cf.SharedItems);
                this.HasSharedItems = true;
            }

            if (cf.FieldGroup != null)
            {
                this.FieldGroup.FromFieldGroup(cf.FieldGroup);
                this.HasFieldGroup = true;
            }

            MemberPropertiesMap mpm;
            using (OpenXmlReader oxr = OpenXmlReader.Create(cf))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(MemberPropertiesMap))
                    {
                        mpm = (MemberPropertiesMap)oxr.LoadCurrentElement();
                        if (mpm.Val != null) this.MemberPropertiesMaps.Add(mpm.Val.Value);
                        else this.MemberPropertiesMaps.Add(0);
                    }
                }
            }
        }

        internal CacheField ToCacheField()
        {
            CacheField cf = new CacheField();
            cf.Name = this.Name;
            if (this.Caption != null && this.Caption.Length > 0) cf.Caption = this.Caption;
            if (this.PropertyName != null & this.PropertyName.Length > 0) cf.PropertyName = this.PropertyName;
            if (this.ServerField != false) cf.ServerField = this.ServerField;
            if (this.UniqueList != true) cf.UniqueList = this.UniqueList;
            if (this.NumberFormatId != null) cf.NumberFormatId = this.NumberFormatId.Value;
            if (this.Formula != null && this.Formula.Length > 0) cf.Formula = this.Formula;
            if (this.SqlType != 0) cf.SqlType = this.SqlType;
            if (this.Hierarchy != 0) cf.Hierarchy = this.Hierarchy;
            if (this.Level != 0) cf.Level = this.Level;
            if (this.DatabaseField != true) cf.DatabaseField = this.DatabaseField;
            if (this.MappingCount != null) cf.MappingCount = this.MappingCount.Value;
            if (this.MemberPropertyField != false) cf.MemberPropertyField = this.MemberPropertyField;

            if (this.HasSharedItems)
            {
                cf.SharedItems = this.SharedItems.ToSharedItems();
            }

            if (this.HasFieldGroup)
            {
                cf.FieldGroup = this.FieldGroup.ToFieldGroup();
            }

            foreach (int i in this.MemberPropertiesMaps)
            {
                if (i != 0) cf.Append(new MemberPropertiesMap() { Val = i });
                else cf.Append(new MemberPropertiesMap());
            }

            return cf;
        }

        internal SLCacheField Clone()
        {
            SLCacheField cf = new SLCacheField();
            cf.Name = this.Name;
            cf.Caption = this.Caption;
            cf.PropertyName = this.PropertyName;
            cf.ServerField = this.ServerField;
            cf.UniqueList = this.UniqueList;
            cf.NumberFormatId = this.NumberFormatId;
            cf.Formula = this.Formula;
            cf.SqlType = this.SqlType;
            cf.Hierarchy = this.Hierarchy;
            cf.Level = this.Level;
            cf.DatabaseField = this.DatabaseField;
            cf.MappingCount = this.MappingCount;
            cf.MemberPropertyField = this.MemberPropertyField;

            cf.HasSharedItems = this.HasSharedItems;
            cf.SharedItems = this.SharedItems.Clone();

            cf.HasFieldGroup = this.HasFieldGroup;
            cf.FieldGroup = this.FieldGroup.Clone();

            cf.MemberPropertiesMaps = new List<int>();
            foreach (int i in this.MemberPropertiesMaps)
            {
                cf.MemberPropertiesMaps.Add(i);
            }

            return cf;
        }
    }
}
