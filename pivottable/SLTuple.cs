using DocumentFormat.OpenXml.Spreadsheet;
using Tuple = DocumentFormat.OpenXml.Spreadsheet.Tuple;

// Apparently .NET Framework 4 has a System.Tuple, which clashes
// with DocumentFormat.OpenXml.Spreadsheet.Tuple.
// Good thing we're on 3.5...

namespace SpreadsheetLight
{
    internal class SLTuple
    {
        internal uint? Field { get; set; }
        internal uint? Hierarchy { get; set; }
        internal uint Item { get; set; }

        internal SLTuple()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Field = null;
            this.Hierarchy = null;
            this.Item = 0;
        }

        internal void FromTuple(Tuple t)
        {
            this.SetAllNull();

            if (t.Field != null) this.Field = t.Field.Value;
            if (t.Hierarchy != null) this.Hierarchy = t.Hierarchy.Value;
            if (t.Item != null) this.Item = t.Item.Value;
        }

        internal Tuple ToTuple()
        {
            Tuple t = new Tuple();
            if (this.Field != null) t.Field = this.Field.Value;
            if (this.Hierarchy != null) t.Hierarchy = this.Hierarchy.Value;
            t.Item = this.Item;

            return t;
        }

        internal SLTuple Clone()
        {
            SLTuple t = new SLTuple();
            t.Field = this.Field;
            t.Hierarchy = this.Hierarchy;
            t.Item = this.Item;

            return t;
        }
    }
}
