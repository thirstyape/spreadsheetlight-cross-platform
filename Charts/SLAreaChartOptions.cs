using System;
using System.Collections.Generic;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Chart customization options for area charts.
    /// </summary>
    public class SLAreaChartOptions
    {
        /// <summary>
        /// Indicates if the area chart has drop lines.
        /// </summary>
        public bool HasDropLines { get; set; }

        /// <summary>
        /// Drop lines properties.
        /// </summary>
        public SLDropLines DropLines { get; set; }

        internal ushort iGapDepth;
        /// <summary>
        /// The gap depth between area clusters (between different data series) as a percentage of width, ranging between 0% and 500% (both inclusive). The default is 150%. This is only used for 3D chart version.
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

        /// <summary>
        /// Initializes an instance of SLAreaChartOptions. It is recommended to use SLChart.CreateAreaChartOptions().
        /// </summary>
        public SLAreaChartOptions()
        {
            this.Initialize(new List<System.Drawing.Color>(), false, false);
        }

        internal SLAreaChartOptions(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.Initialize(ThemeColors, IsStylish, ThrowExceptionsIfAny);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.HasDropLines = false;
            this.DropLines = new SLDropLines(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.iGapDepth = 150;
        }

        /// <summary>
        /// Clone a new instance of SLAreaChartOptions.
        /// </summary>
        /// <returns>An SLAreaChartOptions object.</returns>
        public SLAreaChartOptions Clone()
        {
            SLAreaChartOptions aco = new SLAreaChartOptions();
            aco.HasDropLines = this.HasDropLines;
            aco.DropLines = this.DropLines.Clone();
            aco.iGapDepth = this.iGapDepth;

            return aco;
        }
    }
}
