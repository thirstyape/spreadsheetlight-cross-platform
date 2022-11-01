using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Excel = DocumentFormat.OpenXml.Office.Excel;

namespace SpreadsheetLight
{
    internal class SLConditionalFormattingRule2010
    {
        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformattingrule.aspx

        internal List<Excel.Formula> Formulas { get; set; }

        internal bool HasColorScale;
        internal SLColorScale2010 ColorScale { get; set; }
        internal bool HasDataBar;
        internal SLDataBar2010 DataBar { get; set; }
        internal bool HasIconSet;
        internal SLIconSet2010 IconSet { get; set; }

        internal bool HasDifferentialType;
        internal SLDifferentialFormat DifferentialType { get; set; }

        // extensions (MOAR extensions!?!?)

        internal ConditionalFormatValues? Type { get; set; }

        internal int? Priority { get; set; }
        internal bool StopIfTrue { get; set; }
        internal bool AboveAverage { get; set; }
        internal bool Percent { get; set; }
        internal bool Bottom { get; set; }

        internal bool HasOperator;
        internal ConditionalFormattingOperatorValues Operator { get; set; }

        internal string Text { get; set; }

        internal bool HasTimePeriod;
        internal TimePeriodValues TimePeriod { get; set; }

        internal uint? Rank { get; set; }
        internal int? StandardDeviation { get; set; }
        internal bool EqualAverage { get; set; }
        internal bool ActivePresent { get; set; }
        internal string Id { get; set; }

        internal SLConditionalFormattingRule2010()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Formulas = new List<Excel.Formula>();
            this.ColorScale = new SLColorScale2010();
            this.HasColorScale = false;
            this.DataBar = new SLDataBar2010();
            this.HasDataBar = false;
            this.IconSet = new SLIconSet2010();
            this.HasIconSet = false;
            this.DifferentialType = new SLDifferentialFormat();
            this.HasDifferentialType = false;

            this.Type = null;

            this.Priority = null;
            this.StopIfTrue = false;
            this.AboveAverage = true;
            this.Percent = false;
            this.Bottom = false;
            this.Operator = ConditionalFormattingOperatorValues.Equal;
            this.HasOperator = false;
            this.Text = null;
            this.TimePeriod = TimePeriodValues.Today;
            this.HasTimePeriod = false;
            this.Rank = null;
            this.StandardDeviation = null;
            this.EqualAverage = false;
            this.ActivePresent = false;
            this.Id = null;
        }

        internal void FromConditionalFormattingRule(X14.ConditionalFormattingRule cfr)
        {
            this.SetAllNull();

            if (cfr.Type != null) this.Type = cfr.Type.Value;
            if (cfr.Priority != null) this.Priority = cfr.Priority.Value;
            if (cfr.StopIfTrue != null) this.StopIfTrue = cfr.StopIfTrue.Value;
            if (cfr.AboveAverage != null) this.AboveAverage = cfr.AboveAverage.Value;
            if (cfr.Percent != null) this.Percent = cfr.Percent.Value;
            if (cfr.Bottom != null) this.Bottom = cfr.Bottom.Value;
            if (cfr.Operator != null)
            {
                this.Operator = cfr.Operator.Value;
                this.HasOperator = true;
            }
            if (cfr.Text != null) this.Text = cfr.Text.Value;
            if (cfr.TimePeriod != null)
            {
                this.TimePeriod = cfr.TimePeriod.Value;
                this.HasTimePeriod = true;
            }
            if (cfr.Rank != null) this.Rank = cfr.Rank.Value;
            if (cfr.StandardDeviation != null) this.StandardDeviation = cfr.StandardDeviation.Value;
            if (cfr.EqualAverage != null) this.EqualAverage = cfr.EqualAverage.Value;
            if (cfr.ActivePresent != null) this.ActivePresent = cfr.ActivePresent.Value;
            if (cfr.Id != null) this.Id = cfr.Id.Value;

            using (OpenXmlReader oxr = OpenXmlReader.Create(cfr))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Excel.Formula))
                    {
                        this.Formulas.Add((Excel.Formula)oxr.LoadCurrentElement().CloneNode(true));
                    }
                    else if (oxr.ElementType == typeof(X14.ColorScale))
                    {
                        this.ColorScale = new SLColorScale2010();
                        this.ColorScale.FromColorScale((X14.ColorScale)oxr.LoadCurrentElement());
                        this.HasColorScale = true;
                    }
                    else if (oxr.ElementType == typeof(X14.DataBar))
                    {
                        this.DataBar = new SLDataBar2010();
                        this.DataBar.FromDataBar((X14.DataBar)oxr.LoadCurrentElement());
                        this.HasDataBar = true;
                    }
                    else if (oxr.ElementType == typeof(X14.IconSet))
                    {
                        this.IconSet = new SLIconSet2010();
                        this.IconSet.FromIconSet((X14.IconSet)oxr.LoadCurrentElement());
                        this.HasIconSet = true;
                    }
                    else if (oxr.ElementType == typeof(X14.DifferentialType))
                    {
                        this.DifferentialType = new SLDifferentialFormat();
                        this.DifferentialType.FromDifferentialType((X14.DifferentialType)oxr.LoadCurrentElement());
                        this.HasDifferentialType = true;
                    }
                }
            }
        }

        internal X14.ConditionalFormattingRule ToConditionalFormattingRule()
        {
            X14.ConditionalFormattingRule cfr = new X14.ConditionalFormattingRule();
            if (this.Type != null) cfr.Type = this.Type.Value;
            if (this.Priority != null) cfr.Priority = this.Priority.Value;
            if (this.StopIfTrue) cfr.StopIfTrue = this.StopIfTrue;
            if (!this.AboveAverage) cfr.AboveAverage = this.AboveAverage;
            if (this.Percent) cfr.Percent = this.Percent;
            if (this.Bottom) cfr.Bottom = this.Bottom;
            if (HasOperator) cfr.Operator = this.Operator;
            if (this.Text != null && this.Text.Length > 0) cfr.Text = this.Text;
            if (HasTimePeriod) cfr.TimePeriod = this.TimePeriod;
            if (this.Rank != null) cfr.Rank = this.Rank.Value;
            if (this.StandardDeviation != null) cfr.StandardDeviation = this.StandardDeviation.Value;
            if (this.EqualAverage) cfr.EqualAverage = this.EqualAverage;
            if (this.ActivePresent) cfr.ActivePresent = this.ActivePresent;
            if (this.Id != null) cfr.Id = this.Id;

            foreach (Excel.Formula f in this.Formulas)
            {
                cfr.Append((Excel.Formula)f.CloneNode(true));
            }
            if (HasColorScale) cfr.Append(this.ColorScale.ToColorScale());
            if (HasDataBar) cfr.Append(this.DataBar.ToDataBar(this.Priority != null));
            if (HasIconSet) cfr.Append(this.IconSet.ToIconSet());
            if (HasDifferentialType) cfr.Append(this.DifferentialType.ToDifferentialType());

            return cfr;
        }

        internal SLConditionalFormattingRule2010 Clone()
        {
            SLConditionalFormattingRule2010 cfr = new SLConditionalFormattingRule2010();

            cfr.Formulas = new List<Excel.Formula>();
            for (int i = 0; i < this.Formulas.Count; ++i)
            {
                cfr.Formulas.Add((Excel.Formula)this.Formulas[i].CloneNode(true));
            }

            cfr.HasColorScale = this.HasColorScale;
            cfr.ColorScale = this.ColorScale.Clone();
            cfr.HasDataBar = this.HasDataBar;
            cfr.DataBar = this.DataBar.Clone();
            cfr.HasIconSet = this.HasIconSet;
            cfr.IconSet = this.IconSet.Clone();
            cfr.HasDifferentialType = this.HasDifferentialType;
            cfr.DifferentialType = this.DifferentialType.Clone();

            cfr.Type = this.Type;
            cfr.Priority = this.Priority;
            cfr.StopIfTrue = this.StopIfTrue;
            cfr.AboveAverage = this.AboveAverage;
            cfr.Percent = this.Percent;
            cfr.Bottom = this.Bottom;
            cfr.HasOperator = this.HasOperator;
            cfr.Operator = this.Operator;
            cfr.Text = this.Text;
            cfr.HasTimePeriod = this.HasTimePeriod;
            cfr.TimePeriod = this.TimePeriod;
            cfr.Rank = this.Rank;
            cfr.StandardDeviation = this.StandardDeviation;
            cfr.EqualAverage = this.EqualAverage;
            cfr.ActivePresent = this.ActivePresent;
            cfr.Id = this.Id;

            return cfr;
        }
    }
}
