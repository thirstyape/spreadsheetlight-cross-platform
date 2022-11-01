using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLCalculatedMember
    {
        internal string Name { get; set; }
        internal string Mdx { get; set; }
        internal string MemberName { get; set; }
        internal string Hierarchy { get; set; }
        internal string ParentName { get; set; }
        internal int SolveOrder { get; set; }
        internal bool Set { get; set; }

        internal SLCalculatedMember()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Name = "";
            this.Mdx = "";
            this.MemberName = "";
            this.Hierarchy = "";
            this.ParentName = "";
            this.SolveOrder = 0;
            this.Set = false;
        }

        internal void FromCalculatedMember(CalculatedMember cm)
        {
            this.SetAllNull();

            if (cm.Name != null) this.Name = cm.Name.Value;
            if (cm.Mdx != null) this.Mdx = cm.Mdx.Value;
            if (cm.MemberName != null) this.MemberName = cm.MemberName.Value;
            if (cm.Hierarchy != null) this.Hierarchy = cm.Hierarchy.Value;
            if (cm.ParentName != null) this.ParentName = cm.ParentName.Value;
            if (cm.SolveOrder != null) this.SolveOrder = cm.SolveOrder.Value;
            if (cm.Set != null) this.Set = cm.Set.Value;
        }

        internal CalculatedMember ToCalculatedMember()
        {
            CalculatedMember cm = new CalculatedMember();
            cm.Name = this.Name;
            cm.Mdx = this.Mdx;
            if (this.MemberName != null && this.MemberName.Length > 0) cm.MemberName = this.MemberName;
            if (this.Hierarchy != null && this.Hierarchy.Length > 0) cm.Hierarchy = this.Hierarchy;
            if (this.ParentName != null && this.ParentName.Length > 0) cm.ParentName = this.ParentName;
            if (this.SolveOrder != 0) cm.SolveOrder = this.SolveOrder;
            if (this.Set != false) cm.Set = this.Set;

            return cm;
        }

        internal SLCalculatedMember Clone()
        {
            SLCalculatedMember cm = new SLCalculatedMember();
            cm.Name = this.Name;
            cm.Mdx = this.Mdx;
            cm.MemberName = this.MemberName;
            cm.Hierarchy = this.Hierarchy;
            cm.ParentName = this.ParentName;
            cm.SolveOrder = this.SolveOrder;
            cm.Set = this.Set;

            return cm;
        }
    }
}
