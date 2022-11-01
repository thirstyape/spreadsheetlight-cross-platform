using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLTupleSet
    {
        //CT_Set

        internal List<SLTuplesType> Tuples { get; set; }

        internal bool HasSortByTuple;
        internal SLTuplesType SortByTuple { get; set; }

        // count is for number of Tuples
        internal int MaxRank { get; set; }
        internal string SetDefinition { get; set; }
        internal SortValues SortType { get; set; }
        internal bool QueryFailed { get; set; }

        internal SLTupleSet()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Tuples = new List<SLTuplesType>();

            this.HasSortByTuple = false;
            this.SortByTuple = new SLTuplesType();

            this.MaxRank = 0;
            this.SetDefinition = "";
            this.SortType = SortValues.None;
            this.QueryFailed = false;
        }

        internal void FromTupleSet(TupleSet ts)
        {
            this.SetAllNull();

            if (ts.MaxRank != null) this.MaxRank = ts.MaxRank.Value;
            if (ts.SetDefinition != null) this.SetDefinition = ts.SetDefinition.Value;
            if (ts.SortType != null) this.SortType = ts.SortType.Value;
            if (ts.QueryFailed != null) this.QueryFailed = ts.QueryFailed.Value;

            SLTuplesType tt;
            using (OpenXmlReader oxr = OpenXmlReader.Create(ts))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Tuples))
                    {
                        tt = new SLTuplesType();
                        tt.FromTuples((Tuples)oxr.LoadCurrentElement());
                        this.Tuples.Add(tt);
                    }
                    else if (oxr.ElementType == typeof(SortByTuple))
                    {
                        this.SortByTuple.FromSortByTuple((SortByTuple)oxr.LoadCurrentElement());
                        this.HasSortByTuple = true;
                    }
                }
            }
        }

        internal TupleSet ToTupleSet()
        {
            TupleSet ts = new TupleSet();
            if (this.Tuples.Count > 0) ts.Count = (uint)this.Tuples.Count;
            ts.MaxRank = this.MaxRank;
            ts.SetDefinition = this.SetDefinition;
            if (this.SortType != SortValues.None) ts.SortType = this.SortType;
            if (this.QueryFailed != false) ts.QueryFailed = this.QueryFailed;

            if (this.Tuples.Count > 0)
            {
                foreach (SLTuplesType tt in this.Tuples)
                {
                    ts.Append(tt.ToTuples());
                }
            }

            if (this.HasSortByTuple)
            {
                ts.Append(this.SortByTuple.ToSortByTuple());
            }

            return ts;
        }

        internal SLTupleSet Clone()
        {
            SLTupleSet ts = new SLTupleSet();
            ts.MaxRank = this.MaxRank;
            ts.SetDefinition = this.SetDefinition;
            ts.SortType = this.SortType;
            ts.QueryFailed = this.QueryFailed;

            ts.Tuples = new List<SLTuplesType>();
            foreach (SLTuplesType tt in this.Tuples)
            {
                ts.Tuples.Add(tt.Clone());
            }

            ts.HasSortByTuple = this.HasSortByTuple;
            ts.SortByTuple = this.SortByTuple.Clone();

            return ts;
        }
    }
}
