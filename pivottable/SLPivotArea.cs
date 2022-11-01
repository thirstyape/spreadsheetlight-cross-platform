using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLPivotArea
    {
        List<SLPivotAreaReference> PivotAreaReferences { get; set; }

        internal int? Field { get; set; }
        internal PivotAreaValues Type { get; set; }
        internal bool DataOnly { get; set; }
        internal bool LabelOnly { get; set; }
        internal bool GrandRow { get; set; }
        internal bool GrandColumn { get; set; }
        internal bool CacheIndex { get; set; }
        internal bool Outline { get; set; }
        internal string Offset { get; set; }
        internal bool CollapsedLevelsAreSubtotals { get; set; }
        internal PivotTableAxisValues? Axis { get; set; }
        internal uint? FieldPosition { get; set; }

        internal SLPivotArea()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.PivotAreaReferences = new List<SLPivotAreaReference>();

            this.Field = null;
            this.Type = PivotAreaValues.Normal;
            this.DataOnly = true;
            this.LabelOnly = false;
            this.GrandRow = false;
            this.GrandColumn = false;
            this.CacheIndex = false;
            this.Outline = true;
            this.Offset = "";
            this.CollapsedLevelsAreSubtotals = false;
            this.Axis = null;
            this.FieldPosition = null;
        }

        internal void FromPivotArea(PivotArea pa)
        {
            this.SetAllNull();

            if (pa.Field != null) this.Field = pa.Field.Value;
            if (pa.Type != null) this.Type = pa.Type.Value;
            if (pa.DataOnly != null) this.DataOnly = pa.DataOnly.Value;
            if (pa.LabelOnly != null) this.LabelOnly = pa.LabelOnly.Value;
            if (pa.GrandRow != null) this.GrandRow = pa.GrandRow.Value;
            if (pa.GrandColumn != null) this.GrandColumn = pa.GrandColumn.Value;
            if (pa.CacheIndex != null) this.CacheIndex = pa.CacheIndex.Value;
            if (pa.Outline != null) this.Outline = pa.Outline.Value;
            if (pa.Offset != null) this.Offset = pa.Offset.Value;
            if (pa.CollapsedLevelsAreSubtotals != null) this.CollapsedLevelsAreSubtotals = pa.CollapsedLevelsAreSubtotals.Value;
            if (pa.Axis != null) this.Axis = pa.Axis.Value;
            if (pa.FieldPosition != null) this.FieldPosition = pa.FieldPosition.Value;

            SLPivotAreaReference par;
            using (OpenXmlReader oxr = OpenXmlReader.Create(pa))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(PivotAreaReference))
                    {
                        par = new SLPivotAreaReference();
                        par.FromPivotAreaReference((PivotAreaReference)oxr.LoadCurrentElement());
                        this.PivotAreaReferences.Add(par);
                    }
                }
            }
        }

        internal PivotArea ToPivotArea()
        {
            PivotArea pa = new PivotArea();
            if (this.Field != null) pa.Field = this.Field.Value;
            if (this.Type != PivotAreaValues.Normal) pa.Type = this.Type;
            if (this.DataOnly != true) pa.DataOnly = this.DataOnly;
            if (this.LabelOnly != false) pa.LabelOnly = this.LabelOnly;
            if (this.GrandRow != false) pa.GrandRow = this.GrandRow;
            if (this.GrandColumn != false) pa.GrandColumn = this.GrandColumn;
            if (this.CacheIndex != false) pa.CacheIndex = this.CacheIndex;
            if (this.Outline != true) pa.Outline = this.Outline;
            if (this.Offset != null && this.Offset.Length > 0) pa.Offset = this.Offset;
            if (this.CollapsedLevelsAreSubtotals != false) pa.CollapsedLevelsAreSubtotals = this.CollapsedLevelsAreSubtotals;
            if (this.Axis != null) pa.Axis = this.Axis.Value;
            if (this.FieldPosition != null) pa.FieldPosition = this.FieldPosition.Value;

            if (this.PivotAreaReferences.Count > 0)
            {
                pa.PivotAreaReferences = new PivotAreaReferences();
                foreach (SLPivotAreaReference par in this.PivotAreaReferences)
                {
                    pa.PivotAreaReferences.Append(par.ToPivotAreaReference());
                }
            }

            return pa;
        }

        internal SLPivotArea Clone()
        {
            SLPivotArea pa = new SLPivotArea();
            pa.Field = this.Field;
            pa.Type = this.Type;
            pa.DataOnly = this.DataOnly;
            pa.LabelOnly = this.LabelOnly;
            pa.GrandRow = this.GrandRow;
            pa.GrandColumn = this.GrandColumn;
            pa.CacheIndex = this.CacheIndex;
            pa.Outline = this.Outline;
            pa.Offset = this.Offset;
            pa.CollapsedLevelsAreSubtotals = this.CollapsedLevelsAreSubtotals;
            pa.Axis = this.Axis;
            pa.FieldPosition = this.FieldPosition;

            pa.PivotAreaReferences = new List<SLPivotAreaReference>();
            foreach (SLPivotAreaReference par in this.PivotAreaReferences)
            {
                pa.PivotAreaReferences.Add(par.Clone());
            }

            return pa;
        }
    }
}
