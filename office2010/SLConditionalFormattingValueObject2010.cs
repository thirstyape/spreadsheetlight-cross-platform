using System;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Excel = DocumentFormat.OpenXml.Office.Excel;

namespace SpreadsheetLight
{
    internal class SLConditionalFormattingValueObject2010
    {
        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformattingvalueobject.aspx

        internal string Formula { get; set; }
        internal X14.ConditionalFormattingValueObjectTypeValues Type { get; set; }
        internal bool GreaterThanOrEqual { get; set; }
        
        internal SLConditionalFormattingValueObject2010()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Formula = string.Empty;
            this.Type = X14.ConditionalFormattingValueObjectTypeValues.Percentile;
            this.GreaterThanOrEqual = true;
        }

        internal void FromConditionalFormattingValueObject(X14.ConditionalFormattingValueObject cfvo)
        {
            this.SetAllNull();

            if (cfvo.Formula != null) this.Formula = cfvo.Formula.Text;
            this.Type = cfvo.Type.Value;
            if (cfvo.GreaterThanOrEqual != null) this.GreaterThanOrEqual = cfvo.GreaterThanOrEqual.Value;
        }

        internal X14.ConditionalFormattingValueObject ToConditionalFormattingValueObject()
        {
            X14.ConditionalFormattingValueObject cfvo = new X14.ConditionalFormattingValueObject();

            if (this.Formula.Length > 0)
            {
                if (this.Formula.StartsWith("="))
                {
                    cfvo.Formula = new Excel.Formula(this.Formula.Substring(1));
                }
                else
                {
                    cfvo.Formula = new Excel.Formula(this.Formula);
                }
            }
            cfvo.Type = this.Type;
            if (!this.GreaterThanOrEqual) cfvo.GreaterThanOrEqual = this.GreaterThanOrEqual;

            return cfvo;
        }

        internal SLConditionalFormattingValueObject2010 Clone()
        {
            SLConditionalFormattingValueObject2010 cfvo = new SLConditionalFormattingValueObject2010();
            cfvo.Formula = this.Formula;
            cfvo.Type = this.Type;
            cfvo.GreaterThanOrEqual = this.GreaterThanOrEqual;

            return cfvo;
        }
    }
}
