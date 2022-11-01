using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for up-down bars.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.UpDownBars class.
    /// </summary>
    public class SLUpDownBars
    {
        internal ushort iGapWidth;
        /// <summary>
        /// The gap width between consecutive up-down bars as a percentage of the width of the bar, ranging from 0 to 500 (both inclusive).
        /// </summary>
        public ushort GapWidth
        {
            get { return iGapWidth; }
            set
            {
                iGapWidth = value;
                if (iGapWidth > 500) iGapWidth = 500;
            }
        }

        /// <summary>
        /// The up bars.
        /// </summary>
        public SLUpBars UpBars { get; set; }

        /// <summary>
        /// The down bars.
        /// </summary>
        public SLDownBars DownBars { get; set; }

        internal SLUpDownBars(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.iGapWidth = 150;
            this.UpBars = new SLUpBars(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.DownBars = new SLDownBars(ThemeColors, IsStylish, ThrowExceptionsIfAny);
        }

        internal C.UpDownBars ToUpDownBars(bool IsStylish)
        {
            C.UpDownBars udb = new C.UpDownBars();
            udb.GapWidth = new C.GapWidth() { Val = iGapWidth };
            udb.UpBars = this.UpBars.ToUpBars(IsStylish);
            udb.DownBars = this.DownBars.ToDownBars(IsStylish);

            return udb;
        }

        internal SLUpDownBars Clone()
        {
            // use the upbar as reference
            SLUpDownBars udb = new SLUpDownBars(this.UpBars.ShapeProperties.listThemeColors, false, this.UpBars.ShapeProperties.ThrowExceptionsIfAny);
            udb.iGapWidth = this.iGapWidth;
            udb.UpBars = this.UpBars.Clone();
            udb.DownBars = this.DownBars.Clone();

            return udb;
        }
    }
}
