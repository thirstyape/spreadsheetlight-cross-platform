using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLQuery
    {
        internal bool HasTuples;
        internal SLTuplesType Tuples { get; set; }

        internal string Mdx { get; set; }

        internal SLQuery()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.HasTuples = false;
            this.Tuples = new SLTuplesType();

            this.Mdx = "";
        }

        internal void FromQuery(Query q)
        {
            this.SetAllNull();

            if (q.Mdx != null) this.Mdx = q.Mdx.Value;

            if (q.Tuples != null)
            {
                this.Tuples.FromTuples(q.Tuples);
                this.HasTuples = true;
            }
        }

        internal Query ToQuery()
        {
            Query q = new Query();
            q.Mdx = this.Mdx;

            if (this.HasTuples) q.Tuples = this.Tuples.ToTuples();

            return q;
        }

        internal SLQuery Clone()
        {
            SLQuery q = new SLQuery();
            q.Mdx = this.Mdx;
            q.HasTuples = this.HasTuples;
            q.Tuples = this.Tuples.Clone();

            return q;
        }
    }
}
