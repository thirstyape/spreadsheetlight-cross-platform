using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Tuple = DocumentFormat.OpenXml.Spreadsheet.Tuple;

// Apparently .NET Framework 4 has a System.Tuple, which clashes
// with DocumentFormat.OpenXml.Spreadsheet.Tuple.
// Good thing we're on 3.5...

namespace SpreadsheetLight
{
    /// <summary>
    /// This doubles for SortByTuple and Tuples
    /// </summary>
    internal class SLTuplesType
    {
        internal List<SLTuple> Tuples { get; set; }
        internal uint? MemberNameCount { get; set; }

        internal SLTuplesType()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Tuples = new List<SLTuple>();
            this.MemberNameCount = null;
        }

        internal void FromSortByTuple(SortByTuple sbt)
        {
            this.SetAllNull();

            if (sbt.MemberNameCount != null) this.MemberNameCount = sbt.MemberNameCount.Value;

            SLTuple t;
            using (OpenXmlReader oxr = OpenXmlReader.Create(sbt))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Tuple))
                    {
                        t = new SLTuple();
                        t.FromTuple((Tuple)oxr.LoadCurrentElement());
                        this.Tuples.Add(t);
                    }
                }
            }
        }

        internal void FromTuples(Tuples tpls)
        {
            this.SetAllNull();

            if (tpls.MemberNameCount != null) this.MemberNameCount = tpls.MemberNameCount.Value;

            SLTuple t;
            using (OpenXmlReader oxr = OpenXmlReader.Create(tpls))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Tuple))
                    {
                        t = new SLTuple();
                        t.FromTuple((Tuple)oxr.LoadCurrentElement());
                        this.Tuples.Add(t);
                    }
                }
            }
        }

        internal SortByTuple ToSortByTuple()
        {
            SortByTuple sbt = new SortByTuple();
            if (this.MemberNameCount != null) sbt.MemberNameCount = this.MemberNameCount.Value;

            foreach (SLTuple t in this.Tuples)
            {
                sbt.Append(t.ToTuple());
            }

            return sbt;
        }

        internal Tuples ToTuples()
        {
            Tuples tpls = new Tuples();
            if (this.MemberNameCount != null) tpls.MemberNameCount = this.MemberNameCount.Value;

            foreach (SLTuple t in this.Tuples)
            {
                tpls.Append(t.ToTuple());
            }

            return tpls;
        }

        internal SLTuplesType Clone()
        {
            SLTuplesType tt = new SLTuplesType();
            tt.MemberNameCount = this.MemberNameCount;

            tt.Tuples = new List<SLTuple>();
            foreach (SLTuple t in this.Tuples)
            {
                tt.Tuples.Add(t.Clone());
            }

            return tt;
        }
    }
}
