using System;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    internal class SLConditionalFormattingIcon2010
    {
        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformattingicon.aspx

        internal X14.IconSetTypeValues IconSet { get; set; }
        internal uint IconId { get; set; }

        internal SLConditionalFormattingIcon2010()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.IconSet = X14.IconSetTypeValues.ThreeTrafficLights1;
            this.IconId = 0;
        }

        internal void FromConditionalFormattingIcon(X14.ConditionalFormattingIcon cfi)
        {
            this.SetAllNull();

            if (cfi.IconSet != null) this.IconSet = cfi.IconSet.Value;
            if (cfi.IconId != null) this.IconId = cfi.IconId.Value;
        }

        internal X14.ConditionalFormattingIcon ToConditionalFormattingIcon()
        {
            X14.ConditionalFormattingIcon cfi = new X14.ConditionalFormattingIcon();

            cfi.IconSet = this.IconSet;
            cfi.IconId = this.IconId;

            return cfi;
        }

        internal SLConditionalFormattingIcon2010 Clone()
        {
            SLConditionalFormattingIcon2010 cfi = new SLConditionalFormattingIcon2010();
            cfi.IconSet = this.IconSet;
            cfi.IconId = this.IconId;

            return cfi;
        }
    }
}
