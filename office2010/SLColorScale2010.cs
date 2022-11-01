using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    internal class SLColorScale2010
    {
        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.colorscale.aspx

        internal List<SLConditionalFormattingValueObject2010> Cfvos { get; set; }
        internal List<SLColor> Colors { get; set; }

        internal SLColorScale2010()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Cfvos = new List<SLConditionalFormattingValueObject2010>();
            this.Colors = new List<SLColor>();
        }

        internal void FromColorScale(X14.ColorScale cs)
        {
            this.SetAllNull();

            SLConditionalFormattingValueObject2010 cfvo;
            SLColor clr;
            using (OpenXmlReader oxr = OpenXmlReader.Create(cs))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(X14.ConditionalFormattingValueObject))
                    {
                        cfvo = new SLConditionalFormattingValueObject2010();
                        cfvo.FromConditionalFormattingValueObject((X14.ConditionalFormattingValueObject)oxr.LoadCurrentElement());
                        this.Cfvos.Add(cfvo);
                    }
                    else if (oxr.ElementType == typeof(X14.Color))
                    {
                        clr = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                        clr.FromExcel2010Color((X14.Color)oxr.LoadCurrentElement());
                        this.Colors.Add(clr);
                    }
                }
            }
        }

        internal X14.ColorScale ToColorScale()
        {
            X14.ColorScale cs = new X14.ColorScale();
            foreach (SLConditionalFormattingValueObject2010 cfvo in this.Cfvos)
            {
                cs.Append(cfvo.ToConditionalFormattingValueObject());
            }
            foreach (SLColor clr in this.Colors)
            {
                cs.Append(clr.ToExcel2010Color());
            }

            return cs;
        }

        internal SLColorScale2010 Clone()
        {
            SLColorScale2010 cs = new SLColorScale2010();

            int i;
            cs.Cfvos = new List<SLConditionalFormattingValueObject2010>();
            for (i = 0; i < this.Cfvos.Count; ++i)
            {
                cs.Cfvos.Add(this.Cfvos[i].Clone());
            }

            cs.Colors = new List<SLColor>();
            for (i = 0; i < this.Colors.Count; ++i)
            {
                cs.Colors.Add(this.Colors[i].Clone());
            }

            return cs;
        }
    }
}
