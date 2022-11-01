using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Excel = DocumentFormat.OpenXml.Office.Excel;

namespace SpreadsheetLight
{
    internal class SLConditionalFormatting2010
    {
        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformatting.aspx

        internal List<SLConditionalFormattingRule2010> Rules { get; set; }
        internal List<SLCellPointRange> ReferenceSequence { get; set; }

        // extensions?

        internal bool Pivot { get; set; }

        internal SLConditionalFormatting2010()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Rules = new List<SLConditionalFormattingRule2010>();
            this.ReferenceSequence = new List<SLCellPointRange>();
            this.Pivot = false;
        }

        internal void FromConditionalFormatting(X14.ConditionalFormatting cf)
        {
            this.SetAllNull();

            if (cf.Pivot != null) this.Pivot = cf.Pivot.Value;

            using (OpenXmlReader oxr = OpenXmlReader.Create(cf))
            {
                while (oxr.Read())
                {
                    SLConditionalFormattingRule2010 cfr;
                    if (oxr.ElementType == typeof(X14.ConditionalFormattingRule))
                    {
                        cfr = new SLConditionalFormattingRule2010();
                        cfr.FromConditionalFormattingRule((X14.ConditionalFormattingRule)oxr.LoadCurrentElement());
                        this.Rules.Add(cfr);
                    }
                    else if (oxr.ElementType == typeof(Excel.ReferenceSequence))
                    {
                        Excel.ReferenceSequence refseq = (Excel.ReferenceSequence)oxr.LoadCurrentElement();
                        this.ReferenceSequence = SLTool.TranslateRefSeqToCellPointRange(refseq);
                    }
                }
            }
        }

        internal X14.ConditionalFormatting ToConditionalFormatting()
        {
            X14.ConditionalFormatting cf = new X14.ConditionalFormatting();
            // otherwise xm:f and xm:seqref becomes xne:f and xne:seqref
            cf.AddNamespaceDeclaration("xm", SLConstants.NamespaceXm);
            // how come sparklines don't need explicit namespace declarations?

            if (this.Pivot) cf.Pivot = this.Pivot;

            int i;
            for (i = 0; i < this.Rules.Count; ++i)
            {
                cf.Append(this.Rules[i].ToConditionalFormattingRule());
            }

            if (this.ReferenceSequence.Count > 0)
            {
                cf.Append(new Excel.ReferenceSequence(SLTool.TranslateCellPointRangeToRefSeq(this.ReferenceSequence)));
            }

            return cf;
        }

        internal SLConditionalFormatting2010 Clone()
        {
            SLConditionalFormatting2010 cf = new SLConditionalFormatting2010();

            int i;
            cf.Rules = new List<SLConditionalFormattingRule2010>();
            for (i = 0; i < this.Rules.Count; ++i)
            {
                cf.Rules.Add(this.Rules[i].Clone());
            }

            cf.ReferenceSequence = new List<SLCellPointRange>();
            SLCellPointRange cpr;
            for (i = 0; i < this.ReferenceSequence.Count; ++i)
            {
                cpr = new SLCellPointRange(this.ReferenceSequence[i].StartRowIndex, this.ReferenceSequence[i].StartColumnIndex, this.ReferenceSequence[i].EndRowIndex, this.ReferenceSequence[i].EndColumnIndex);
                cf.ReferenceSequence.Add(cpr);
            }

            cf.Pivot = this.Pivot;

            return cf;
        }
    }
}
