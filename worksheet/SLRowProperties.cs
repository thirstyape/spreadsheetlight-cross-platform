using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLRowProperties
    {
        internal bool IsEmpty
        {
            get
            {
                return (this.StyleIndex == 0) && !this.HasHeight && !this.Hidden
                    && (this.OutlineLevel == 0) && !this.Collapsed
                    && !this.ThickTop && !this.ThickBottom && !this.ShowPhonetic;
            }
        }

        internal double DefaultRowHeight { get; set; }

        internal uint StyleIndex { get; set; }

        internal bool CustomHeight;

        internal bool HasHeight;
        private double fHeight;
        // The row height in points.
        internal double Height
        {
            get { return fHeight; }
            set
            {
                double fModifiedRowHeight = value / SLDocument.RowHeightMultiple;
                fModifiedRowHeight = Math.Ceiling(fModifiedRowHeight) * SLDocument.RowHeightMultiple;

                lHeightInEMU = (long)(fModifiedRowHeight * SLConstants.PointToEMU);

                fHeight = fModifiedRowHeight;
                CustomHeight = true;
                HasHeight = true;
            }
        }

        private long lHeightInEMU;
        internal long HeightInEMU
        {
            get { return lHeightInEMU; }
        }

        internal bool Hidden { get; set; }
        internal byte OutlineLevel { get; set; }
        internal bool Collapsed { get; set; }
        internal bool ThickTop { get; set; }
        internal bool ThickBottom { get; set; }
        internal bool ShowPhonetic { get; set; }

        internal SLRowProperties(double DefaultRowHeight)
        {
            this.DefaultRowHeight = DefaultRowHeight;
            this.StyleIndex = 0;
            this.Height = DefaultRowHeight;
            this.CustomHeight = false;
            this.HasHeight = false;
            this.Hidden = false;
            this.OutlineLevel = 0;
            this.Collapsed = false;
            this.ThickTop = false;
            this.ThickBottom = false;
            this.ShowPhonetic = false;
        }

        internal Row ToRow()
        {
            Row r = new Row();
            if (this.StyleIndex > 0)
            {
                r.StyleIndex = this.StyleIndex;
                r.CustomFormat = true;
            }
            if (HasHeight)
            {
                r.Height = this.Height;
            }
            if (CustomHeight)
            {
                r.CustomHeight = true;
            }

            if (this.Hidden != false) r.Hidden = this.Hidden;
            if (this.OutlineLevel > 0) r.OutlineLevel = this.OutlineLevel;
            if (this.Collapsed != false) r.Collapsed = this.Collapsed;
            if (this.ThickTop != false) r.ThickTop = this.ThickTop;
            if (this.ThickBottom != false) r.ThickBot = this.ThickBottom;
            if (this.ShowPhonetic != false) r.ShowPhonetic = this.ShowPhonetic;

            return r;
        }

        internal SLRowProperties Clone()
        {
            SLRowProperties rp = new SLRowProperties(this.DefaultRowHeight);
            rp.DefaultRowHeight = this.DefaultRowHeight;
            rp.StyleIndex = this.StyleIndex;
            rp.HasHeight = this.HasHeight;
            rp.fHeight = this.fHeight;
            rp.lHeightInEMU = this.lHeightInEMU;
            rp.Hidden = this.Hidden;
            rp.OutlineLevel = this.OutlineLevel;
            rp.Collapsed = this.Collapsed;
            rp.ThickTop = this.ThickTop;
            rp.ThickBottom = this.ThickBottom;
            rp.ShowPhonetic = this.ShowPhonetic;

            return rp;
        }
    }
}
