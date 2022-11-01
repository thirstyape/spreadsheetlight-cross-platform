using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLMembers
    {
        internal List<string> Members { get; set; }
        internal uint? Level { get; set; }

        internal SLMembers()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Members = new List<string>();
            this.Level = null;
        }

        internal void FromMembers(Members m)
        {
            this.SetAllNull();

            if (m.Level != null) this.Level = m.Level.Value;

            Member mem;
            using (OpenXmlReader oxr = OpenXmlReader.Create(m))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Member))
                    {
                        mem = (Member)oxr.LoadCurrentElement();
                        this.Members.Add(mem.Name.Value);
                    }
                }
            }
        }

        internal Members ToMembers()
        {
            Members m = new Members();
            m.Count = (uint)this.Members.Count;
            if (this.Level != null) m.Level = this.Level.Value;

            foreach (string s in this.Members)
            {
                m.Append(new Member() { Name = s });
            }

            return m;
        }

        internal SLMembers Clone()
        {
            SLMembers m = new SLMembers();
            m.Level = this.Level;

            m.Members = new List<string>();
            foreach (string s in this.Members)
            {
                m.Members.Add(s);
            }

            return m;
        }
    }
}
