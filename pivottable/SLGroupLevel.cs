using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLGroupLevel
    {
        internal List<SLGroup> Groups { get; set; }

        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal bool User { get; set; }
        internal bool CustomRollUp { get; set; }

        internal SLGroupLevel()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Groups = new List<SLGroup>();

            this.UniqueName = "";
            this.Caption = "";
            this.User = false;
            this.CustomRollUp = false;
        }

        internal void FromGroupLevel(GroupLevel gl)
        {
            this.SetAllNull();

            if (gl.UniqueName != null) this.UniqueName = gl.UniqueName.Value;
            if (gl.Caption != null) this.Caption = gl.Caption.Value;
            if (gl.User != null) this.User = gl.User.Value;
            if (gl.CustomRollUp != null) this.CustomRollUp = gl.CustomRollUp.Value;

            if (gl.Groups != null)
            {
                SLGroup g;
                using (OpenXmlReader oxr = OpenXmlReader.Create(gl.Groups))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(Group))
                        {
                            g = new SLGroup();
                            g.FromGroup((Group)oxr.LoadCurrentElement());
                            this.Groups.Add(g);
                        }
                    }
                }
            }
        }

        internal GroupLevel ToGroupLevel()
        {
            GroupLevel gl = new GroupLevel();
            gl.UniqueName = this.UniqueName;
            gl.Caption = this.Caption;
            if (this.User != false) gl.User = this.User;
            if (this.CustomRollUp != false) gl.CustomRollUp = this.CustomRollUp;

            if (this.Groups.Count > 0)
            {
                gl.Groups = new Groups() { Count = (uint)this.Groups.Count };
                foreach (SLGroup g in this.Groups)
                {
                    gl.Groups.Append(g.ToGroup());
                }
            }

            return gl;
        }

        internal SLGroupLevel Clone()
        {
            SLGroupLevel gl = new SLGroupLevel();
            gl.UniqueName = this.UniqueName;
            gl.Caption = this.Caption;
            gl.User = this.User;
            gl.CustomRollUp = this.CustomRollUp;

            gl.Groups = new List<SLGroup>();
            foreach (SLGroup g in this.Groups)
            {
                gl.Groups.Add(g.Clone());
            }

            return gl;
        }
    }
}
