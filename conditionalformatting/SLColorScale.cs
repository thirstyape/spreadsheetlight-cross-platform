using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLColorScale
    {
        internal List<SLConditionalFormatValueObject> Cfvos { get; set; }
        internal List<SLColor> Colors { get; set; }

        internal SLColorScale()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Cfvos = new List<SLConditionalFormatValueObject>();
            this.Colors = new List<SLColor>();
        }

        internal void FromColorScale(ColorScale cs)
        {
            this.SetAllNull();

            SLConditionalFormatValueObject cfvo;
            SLColor clr;
            using (OpenXmlReader oxr = OpenXmlReader.Create(cs))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(ConditionalFormatValueObject))
                    {
                        cfvo = new SLConditionalFormatValueObject();
                        cfvo.FromConditionalFormatValueObject((ConditionalFormatValueObject)oxr.LoadCurrentElement());
                        this.Cfvos.Add(cfvo);
                    }
                    else if (oxr.ElementType == typeof(Color))
                    {
                        clr = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                        clr.FromSpreadsheetColor((Color)oxr.LoadCurrentElement());
                        this.Colors.Add(clr);
                    }
                }
            }
        }

        internal ColorScale ToColorScale()
        {
            ColorScale cs = new ColorScale();
            foreach (SLConditionalFormatValueObject cfvo in this.Cfvos)
            {
                cs.Append(cfvo.ToConditionalFormatValueObject());
            }
            foreach (SLColor clr in this.Colors)
            {
                cs.Append(clr.ToSpreadsheetColor());
            }

            return cs;
        }

        internal SLColorScale2010 ToSLColorScale2010()
        {
            SLColorScale2010 cs2010 = new SLColorScale2010();
            foreach (SLConditionalFormatValueObject cfvo in this.Cfvos)
            {
                cs2010.Cfvos.Add(cfvo.ToSLConditionalFormattingValueObject2010());
            }
            foreach (SLColor clr in this.Colors)
            {
                cs2010.Colors.Add(clr.Clone());
            }

            return cs2010;
        }

        internal SLColorScale Clone()
        {
            SLColorScale cs = new SLColorScale();

            int i;
            cs.Cfvos = new List<SLConditionalFormatValueObject>();
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
