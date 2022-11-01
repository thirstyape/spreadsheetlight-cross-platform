using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLConditionalFormat
    {
        internal List<SLPivotArea> PivotAreas { get; set; }
        internal ScopeValues Scope { get; set; }
        internal RuleValues Type { get; set; }
        internal uint Priority { get; set; }

        internal SLConditionalFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.PivotAreas = new List<SLPivotArea>();
            this.Scope = ScopeValues.Selection;
            this.Type = RuleValues.None;
            this.Priority = 0;
        }

        internal void FromConditionalFormat(ConditionalFormat cf)
        {
            this.SetAllNull();

            if (cf.Scope != null) this.Scope = cf.Scope.Value;
            if (cf.Type != null) this.Type = cf.Type.Value;
            if (cf.Priority != null) this.Priority = cf.Priority.Value;

            SLPivotArea pa;
            using (OpenXmlReader oxr = OpenXmlReader.Create(cf))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(PivotArea))
                    {
                        pa = new SLPivotArea();
                        pa.FromPivotArea((PivotArea)oxr.LoadCurrentElement());
                        this.PivotAreas.Add(pa);
                    }
                }
            }
        }

        internal ConditionalFormat ToConditionalFormat()
        {
            ConditionalFormat cf = new ConditionalFormat();
            cf.PivotAreas = new PivotAreas() { Count = (uint)this.PivotAreas.Count };
            foreach (SLPivotArea pa in this.PivotAreas)
            {
                cf.PivotAreas.Append(pa.ToPivotArea());
            }

            if (this.Scope != ScopeValues.Selection) cf.Scope = this.Scope;
            if (this.Type != RuleValues.None) cf.Type = this.Type;
            cf.Priority = this.Priority;

            return cf;
        }

        internal SLConditionalFormat Clone()
        {
            SLConditionalFormat cf = new SLConditionalFormat();
            cf.Scope = this.Scope;
            cf.Type = this.Type;
            cf.Priority = this.Priority;

            cf.PivotAreas = new List<SLPivotArea>();
            foreach (SLPivotArea pa in this.PivotAreas)
            {
                cf.PivotAreas.Add(pa.Clone());
            }

            return cf;
        }
    }
}
