using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLGroup
    {
        internal List<SLGroupMember> GroupMembers { get; set; }

        internal string Name { get; set; }
        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal string UniqueParent { get; set; }
        internal int? Id { get; set; }

        internal SLGroup()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.GroupMembers = new List<SLGroupMember>();

            this.Name = "";
            this.UniqueName = "";
            this.Caption = "";
            this.UniqueParent = "";
            this.Id = null;
        }

        internal void FromGroup(Group g)
        {
            this.SetAllNull();

            if (g.Name != null) this.Name = g.Name.Value;
            if (g.UniqueName != null) this.UniqueName = g.UniqueName.Value;
            if (g.Caption != null) this.Caption = g.Caption.Value;
            if (g.UniqueParent != null) this.UniqueParent = g.UniqueParent.Value;
            if (g.Id != null) this.Id = g.Id.Value;
        }

        internal Group ToGroup()
        {
            Group g = new Group();
            g.Name = this.Name;
            g.UniqueName = this.UniqueName;
            g.Caption = this.Caption;
            if (this.UniqueParent != null && this.UniqueParent.Length > 0) g.UniqueParent = this.UniqueParent;
            if (this.Id != null) g.Id = this.Id.Value;

            if (this.GroupMembers.Count > 0)
            {
                g.GroupMembers = new GroupMembers() { Count = (uint)this.GroupMembers.Count };
                foreach (SLGroupMember gm in this.GroupMembers)
                {
                    g.GroupMembers.Append(gm.ToGroupMember());
                }
            }

            return g;
        }

        internal SLGroup Clone()
        {
            SLGroup g = new SLGroup();
            g.Name = this.Name;
            g.UniqueName = this.UniqueName;
            g.Caption = this.Caption;
            g.UniqueParent = this.UniqueParent;
            g.Id = this.Id;

            g.GroupMembers = new List<SLGroupMember>();
            foreach (SLGroupMember gm in this.GroupMembers)
            {
                g.GroupMembers.Add(gm.Clone());
            }

            return g;
        }
    }
}
