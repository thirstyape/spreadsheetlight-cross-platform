using System;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Chart customization options for bar and column charts.
    /// </summary>
    public class SLBarChartOptions
    {
        internal ushort iGapWidth;
        /// <summary>
        /// The gap width between bar or column clusters (in the same data series) as a percentage of bar or column width, ranging between 0% and 500% (both inclusive). The default is 150%.
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

        internal ushort iGapDepth;
        /// <summary>
        /// The gap depth between bar or columns clusters (between different data series) as a percentage of bar or column width, ranging between 0% and 500% (both inclusive). The default is 150%. This is only used for 3D chart version.
        /// </summary>
        public ushort GapDepth
        {
            get { return iGapDepth; }
            set
            {
                iGapDepth = value;
                if (iGapDepth > 500) iGapDepth = 500;
            }
        }

        internal sbyte byOverlap;
        /// <summary>
        /// The amount of overlapping for bars and columns on 2D bar/column charts, ranging from -100 to 100 (both inclusive). The default is 0. For stacked and "100% stacked" bar/column charts, this should be 100.
        /// </summary>
        public sbyte Overlap
        {
            get { return byOverlap; }
            set
            {
                byOverlap = value;
                if (byOverlap < -100) byOverlap = -100;
                if (byOverlap > 100) byOverlap = 100;
            }
        }

        /// <summary>
        /// Initializes an instance of SLBarChartOptions.
        /// </summary>
        public SLBarChartOptions()
        {
            this.iGapWidth = 150;
            this.iGapDepth = 150;
            this.byOverlap = 0;
        }

        internal SLBarChartOptions Clone()
        {
            SLBarChartOptions bco = new SLBarChartOptions();
            bco.iGapWidth = this.iGapWidth;
            bco.iGapDepth = this.iGapDepth;
            bco.byOverlap = this.byOverlap;

            return bco;
        }
    }
}
