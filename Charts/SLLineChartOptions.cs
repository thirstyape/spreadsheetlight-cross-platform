using System;
using System.Collections.Generic;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Chart customization options for line charts.
    /// </summary>
    public class SLLineChartOptions
    {
        internal ushort iGapDepth;
        /// <summary>
        /// The gap depth between line clusters (between different data series) as a percentage of bar or column width, ranging between 0% and 500% (both inclusive). The default is 150%. This is only used for 3D chart version.
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
        /// Indicates if the line chart has drop lines.
        /// </summary>
        public bool HasDropLines { get; set; }

        /// <summary>
        /// Drop lines properties.
        /// </summary>
        public SLDropLines DropLines { get; set; }

        /// <summary>
        /// Indicates if the line chart has high-low lines. This is not applicable for 3D line charts.
        /// </summary>
        public bool HasHighLowLines { get; set; }

        /// <summary>
        /// High-low lines properties.
        /// </summary>
        public SLHighLowLines HighLowLines { get; set; }

        /// <summary>
        /// Indicates if the line chart has up-down bars. This is not applicable for 3D line charts.
        /// </summary>
        public bool HasUpDownBars { get; set; }

        /// <summary>
        /// Up-down bars properties.
        /// </summary>
        public SLUpDownBars UpDownBars { get; set; }

        /// <summary>
        /// Whether the line connecting data points use C splines (instead of straight lines).
        /// </summary>
        public bool Smooth { get; set; }

        /// <summary>
        /// Initializes an instance of SLLineChartOptions. It is recommended to use SLChart.CreateLineChartOptions().
        /// </summary>
        public SLLineChartOptions()
        {
            this.Initialize(new List<System.Drawing.Color>(), false, false);
        }

        internal SLLineChartOptions(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.Initialize(ThemeColors, IsStylish, ThrowExceptionsIfAny);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.iGapDepth = 150;
            this.HasDropLines = false;
            this.DropLines = new SLDropLines(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.HasHighLowLines = false;
            this.HighLowLines = new SLHighLowLines(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.HasUpDownBars = false;
            this.UpDownBars = new SLUpDownBars(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.Smooth = false;
        }

        /// <summary>
        /// Clone an instance of SLLineChartOptions.
        /// </summary>
        /// <returns>An SLLineChartOptions object.</returns>
        public SLLineChartOptions Clone()
        {
            SLLineChartOptions lco = new SLLineChartOptions();
            lco.iGapDepth = this.iGapDepth;
            lco.HasDropLines = this.HasDropLines;
            lco.DropLines = this.DropLines.Clone();
            lco.HasHighLowLines = this.HasHighLowLines;
            lco.HighLowLines = this.HighLowLines.Clone();
            lco.HasUpDownBars = this.HasUpDownBars;
            lco.UpDownBars = this.UpDownBars.Clone();
            lco.Smooth = this.Smooth;

            return lco;
        }
    }
}
