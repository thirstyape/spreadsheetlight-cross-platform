using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLCacheHierarchy
    {
        internal List<int> FieldsUsage { get; set; }
        internal List<SLGroupLevel> GroupLevels { get; set; }

        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal bool Measure { get; set; }
        internal bool Set { get; set; }
        internal uint? ParentSet { get; set; }
        internal int IconSet { get; set; }
        internal bool Attribute { get; set; }
        internal bool Time { get; set; }
        internal bool KeyAttribute { get; set; }
        internal string DefaultMemberUniqueName { get; set; }
        internal string AllUniqueName { get; set; }
        internal string AllCaption { get; set; }
        internal string DimensionUniqueName { get; set; }
        internal string DisplayFolder { get; set; }
        internal string MeasureGroup { get; set; }
        internal bool Measures { get; set; }
        internal uint Count { get; set; }
        internal bool OneField { get; set; }
        internal ushort? MemberValueDatatype { get; set; }
        internal bool? Unbalanced { get; set; }
        internal bool? UnbalancedGroup { get; set; }
        internal bool Hidden { get; set; }

        internal SLCacheHierarchy()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.FieldsUsage = new List<int>();
            this.GroupLevels = new List<SLGroupLevel>();

            this.UniqueName = "";
            this.Caption = "";
            this.Measure = false;
            this.Set = false;
            this.ParentSet = null;
            this.IconSet = 0;
            this.Attribute = false;
            this.Time = false;
            this.KeyAttribute = false;
            this.DefaultMemberUniqueName = "";
            this.AllUniqueName = "";
            this.AllCaption = "";
            this.DimensionUniqueName = "";
            this.DisplayFolder = "";
            this.MeasureGroup = "";
            this.Measures = false;
            this.Count = 0;
            this.OneField = false;
            this.MemberValueDatatype = null;
            this.Unbalanced = null;
            this.UnbalancedGroup = null;
            this.Hidden = false;
        }

        internal void FromCacheHierarchy(CacheHierarchy ch)
        {
            this.SetAllNull();

            if (ch.UniqueName != null) this.UniqueName = ch.UniqueName.Value;
            if (ch.Caption != null) this.Caption = ch.Caption.Value;
            if (ch.Measure != null) this.Measure = ch.Measure.Value;
            if (ch.Set != null) this.Set = ch.Set.Value;
            if (ch.ParentSet != null) this.ParentSet = ch.ParentSet.Value;
            if (ch.IconSet != null) this.IconSet = ch.IconSet.Value;
            if (ch.Attribute != null) this.Attribute = ch.Attribute.Value;
            if (ch.Time != null) this.Time = ch.Time.Value;
            if (ch.KeyAttribute != null) this.KeyAttribute = ch.KeyAttribute.Value;
            if (ch.DefaultMemberUniqueName != null) this.DefaultMemberUniqueName = ch.DefaultMemberUniqueName.Value;
            if (ch.AllUniqueName != null) this.AllUniqueName = ch.AllUniqueName.Value;
            if (ch.AllCaption != null) this.AllCaption = ch.AllCaption.Value;
            if (ch.DimensionUniqueName != null) this.DimensionUniqueName = ch.DimensionUniqueName.Value;
            if (ch.DisplayFolder != null) this.DisplayFolder = ch.DisplayFolder.Value;
            if (ch.MeasureGroup != null) this.MeasureGroup = ch.MeasureGroup.Value;
            if (ch.Measures != null) this.Measures = ch.Measures.Value;
            if (ch.Count != null) this.Count = ch.Count.Value;
            if (ch.OneField != null) this.OneField = ch.OneField.Value;
            if (ch.MemberValueDatatype != null) this.MemberValueDatatype = ch.MemberValueDatatype.Value;
            if (ch.Unbalanced != null) this.Unbalanced = ch.Unbalanced.Value;
            if (ch.UnbalancedGroup != null) this.UnbalancedGroup = ch.UnbalancedGroup.Value;
            if (ch.Hidden != null) this.Hidden = ch.Hidden.Value;

            FieldUsage fu;
            SLGroupLevel gl;
            using (OpenXmlReader oxr = OpenXmlReader.Create(ch))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(FieldUsage))
                    {
                        fu = (FieldUsage)oxr.LoadCurrentElement();
                        this.FieldsUsage.Add(fu.Index.Value);
                    }
                    else if (oxr.ElementType == typeof(GroupLevel))
                    {
                        gl = new SLGroupLevel();
                        gl.FromGroupLevel((GroupLevel)oxr.LoadCurrentElement());
                        this.GroupLevels.Add(gl);
                    }
                }
            }
        }

        internal CacheHierarchy ToCacheHierarchy()
        {
            CacheHierarchy ch = new CacheHierarchy();
            ch.UniqueName = this.UniqueName;
            if (this.Caption != null && this.Caption.Length > 0) ch.Caption = this.Caption;
            if (this.Measure != false) ch.Measure = this.Measure;
            if (this.Set != false) ch.Set = this.Set;
            if (this.ParentSet != null) ch.ParentSet = this.ParentSet.Value;
            if (this.IconSet != 0) ch.IconSet = this.IconSet;
            if (this.Attribute != false) ch.Attribute = this.Attribute;
            if (this.Time != false) ch.Time = this.Time;
            if (this.KeyAttribute != false) ch.KeyAttribute = this.KeyAttribute;
            if (this.DefaultMemberUniqueName != null && this.DefaultMemberUniqueName.Length > 0) ch.DefaultMemberUniqueName = this.DefaultMemberUniqueName;
            if (this.AllUniqueName != null && this.AllUniqueName.Length > 0) ch.AllUniqueName = this.AllUniqueName;
            if (this.AllCaption != null && this.AllCaption.Length > 0) ch.AllCaption = this.AllCaption;
            if (this.DimensionUniqueName != null && this.DimensionUniqueName.Length > 0) ch.DimensionUniqueName = this.DimensionUniqueName;
            if (this.DisplayFolder != null && this.DisplayFolder.Length > 0) ch.DisplayFolder = this.DisplayFolder;
            if (this.MeasureGroup != null && this.MeasureGroup.Length > 0) ch.MeasureGroup = this.MeasureGroup;
            if (this.Measures != false) ch.Measures = this.Measures;
            ch.Count = this.Count;
            if (this.OneField != false) ch.OneField = this.OneField;
            if (this.MemberValueDatatype != null) ch.MemberValueDatatype = this.MemberValueDatatype.Value;
            if (this.Unbalanced != null) ch.Unbalanced = this.Unbalanced.Value;
            if (this.UnbalancedGroup != null) ch.UnbalancedGroup = this.UnbalancedGroup.Value;
            if (this.Hidden != false) ch.Hidden = this.Hidden;

            if (this.FieldsUsage.Count > 0)
            {
                ch.FieldsUsage = new FieldsUsage() { Count = (uint)this.FieldsUsage.Count };
                foreach (int i in this.FieldsUsage)
                {
                    ch.FieldsUsage.Append(new FieldUsage() { Index = i });
                }
            }

            if (this.GroupLevels.Count > 0)
            {
                ch.GroupLevels = new GroupLevels() { Count = (uint)this.GroupLevels.Count };
                foreach (SLGroupLevel gl in this.GroupLevels)
                {
                    ch.GroupLevels.Append(gl.ToGroupLevel());
                }
            }

            return ch;
        }

        internal SLCacheHierarchy Clone()
        {
            SLCacheHierarchy ch = new SLCacheHierarchy();
            ch.UniqueName = this.UniqueName;
            ch.Caption = this.Caption;
            ch.Measure = this.Measure;
            ch.Set = this.Set;
            ch.ParentSet = this.ParentSet;
            ch.IconSet = this.IconSet;
            ch.Attribute = this.Attribute;
            ch.Time = this.Time;
            ch.KeyAttribute = this.KeyAttribute;
            ch.DefaultMemberUniqueName = this.DefaultMemberUniqueName;
            ch.AllUniqueName = this.AllUniqueName;
            ch.AllCaption = this.AllCaption;
            ch.DimensionUniqueName = this.DimensionUniqueName;
            ch.DisplayFolder = this.DisplayFolder;
            ch.MeasureGroup = this.MeasureGroup;
            ch.Measures = this.Measures;
            ch.Count = this.Count;
            ch.OneField = this.OneField;
            ch.MemberValueDatatype = this.MemberValueDatatype;
            ch.Unbalanced = this.Unbalanced;
            ch.UnbalancedGroup = this.UnbalancedGroup;
            ch.Hidden = this.Hidden;

            ch.FieldsUsage = new List<int>();
            foreach (int i in this.FieldsUsage)
            {
                ch.FieldsUsage.Add(i);
            }

            ch.GroupLevels = new List<SLGroupLevel>();
            foreach (SLGroupLevel gl in this.GroupLevels)
            {
                ch.GroupLevels.Add(gl.Clone());
            }

            return ch;
        }
    }
}
