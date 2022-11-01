using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    class SLIconSet2010
    {
        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.iconset.aspx

        internal List<SLConditionalFormattingValueObject2010> Cfvos { get; set; }
        internal List<SLConditionalFormattingIcon2010> CustomIcons { get; set; }
        internal X14.IconSetTypeValues IconSetType { get; set; }
        internal bool ShowValue { get; set; }
        internal bool Percent { get; set; }
        internal bool Reverse { get; set; }
        
        // This is true if and only if CustomIcons is used.
        // So we'll just ignore it and focus on the number of CustomIcons instead.
        // internal bool Custom

        internal SLIconSet2010()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Cfvos = new List<SLConditionalFormattingValueObject2010>();
            this.CustomIcons = new List<SLConditionalFormattingIcon2010>();
            this.IconSetType = X14.IconSetTypeValues.ThreeTrafficLights1;
            this.ShowValue = true;
            this.Percent = true;
            this.Reverse = false;
        }

        internal void FromIconSet(X14.IconSet ics)
        {
            this.SetAllNull();

            if (ics.IconSetTypes != null) this.IconSetType = ics.IconSetTypes.Value;
            if (ics.ShowValue != null) this.ShowValue = ics.ShowValue.Value;
            if (ics.Percent != null) this.Percent = ics.Percent.Value;
            if (ics.Reverse != null) this.Reverse = ics.Reverse.Value;

            using (OpenXmlReader oxr = OpenXmlReader.Create(ics))
            {
                SLConditionalFormattingValueObject2010 cfvo;
                SLConditionalFormattingIcon2010 cfi;
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(X14.ConditionalFormattingValueObject))
                    {
                        cfvo = new SLConditionalFormattingValueObject2010();
                        cfvo.FromConditionalFormattingValueObject((X14.ConditionalFormattingValueObject)oxr.LoadCurrentElement());
                        this.Cfvos.Add(cfvo);
                    }
                    else if (oxr.ElementType == typeof(X14.ConditionalFormattingIcon))
                    {
                        cfi = new SLConditionalFormattingIcon2010();
                        cfi.FromConditionalFormattingIcon((X14.ConditionalFormattingIcon)oxr.LoadCurrentElement());
                        this.CustomIcons.Add(cfi);
                    }
                }
            }
        }

        internal X14.IconSet ToIconSet()
        {
            X14.IconSet ics = new X14.IconSet();
            if (this.IconSetType != X14.IconSetTypeValues.ThreeTrafficLights1) ics.IconSetTypes = this.IconSetType;
            if (!this.ShowValue) ics.ShowValue = this.ShowValue;
            if (!this.Percent) ics.Percent = this.Percent;
            if (this.Reverse) ics.Reverse = this.Reverse;
            if (this.CustomIcons.Count > 0) ics.Custom = true;

            foreach (SLConditionalFormattingValueObject2010 cfvo in this.Cfvos)
            {
                ics.Append(cfvo.ToConditionalFormattingValueObject());
            }

            foreach (SLConditionalFormattingIcon2010 cfi in this.CustomIcons)
            {
                ics.Append(cfi.ToConditionalFormattingIcon());
            }

            return ics;
        }

        internal SLIconSet2010 Clone()
        {
            SLIconSet2010 ics = new SLIconSet2010();

            int i;

            ics.Cfvos = new List<SLConditionalFormattingValueObject2010>();
            for (i = 0; i < this.Cfvos.Count; ++i)
            {
                ics.Cfvos.Add(this.Cfvos[i].Clone());
            }

            ics.CustomIcons = new List<SLConditionalFormattingIcon2010>();
            for (i = 0; i < this.CustomIcons.Count; ++i)
            {
                ics.CustomIcons.Add(this.CustomIcons[i].Clone());
            }

            ics.IconSetType = this.IconSetType;
            ics.ShowValue = this.ShowValue;
            ics.Percent = this.Percent;
            ics.Reverse = this.Reverse;

            return ics;
        }
    }
}
