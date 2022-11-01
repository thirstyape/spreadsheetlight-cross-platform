using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLGroupMember
    {
        internal string UniqueName { get; set; }
        internal bool Group { get; set; }

        internal SLGroupMember()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.UniqueName = "";
            this.Group = false;
        }

        internal void FromGroupMember(GroupMember gm)
        {
            this.SetAllNull();

            if (gm.UniqueName != null) this.UniqueName = gm.UniqueName.Value;
            if (gm.Group != null) this.Group = gm.Group.Value;
        }

        internal GroupMember ToGroupMember()
        {
            GroupMember gm = new GroupMember();
            gm.UniqueName = this.UniqueName;
            if (this.Group!=false)gm.Group=this.Group;

            return gm;
        }

        internal SLGroupMember Clone()
        {
            SLGroupMember gm = new SLGroupMember();
            gm.UniqueName = this.UniqueName;
            gm.Group = this.Group;

            return gm;
        }
    }
}
