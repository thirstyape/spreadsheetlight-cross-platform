using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Chart customization options for pie, bar-of-pie, pie-of-pie and doughnut charts.
    /// </summary>
    public class SLPieChartOptions
    {
        /// <summary>
        /// Each data point shall have a different color. The default is "true".
        /// </summary>
        public bool VaryColors { get; set; }

        internal ushort iFirstSliceAngle;
        /// <summary>
        /// Angle of the first slice, ranging from 0 degrees to 360 degrees.
        /// </summary>
        public ushort FirstSliceAngle
        {
            get { return iFirstSliceAngle; }
            set
            {
                iFirstSliceAngle = value;
                if (iFirstSliceAngle > 360) iFirstSliceAngle = 360;
            }
        }

        internal byte byHoleSize;
        /// <summary>
        /// The size of the hole in a doughnut chart, ranging from 10% to 90% of the diameter of the doughnut chart. If the doughnut chart is exploded, the diameter is taken to be that when it's not exploded.
        /// </summary>
        public byte HoleSize
        {
            get { return byHoleSize; }
            set
            {
                byHoleSize = value;
                if (byHoleSize < 10) byHoleSize = 10;
                if (byHoleSize > 90) byHoleSize = 90;
            }
        }

        internal ushort iGapWidth;
        /// <summary>
        /// The gap width between the first pie and the second bar or pie chart, ranging from 0 to 500 (both inclusive). This is for bar-of-pie or pie-of-pie charts.
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

        internal bool HasSplit;
        internal C.SplitValues SplitType { get; set; }
        internal double SplitPosition { get; set; }
        internal List<int> SecondPiePoints { get; set; }

        internal ushort iSecondPieSize;
        /// <summary>
        /// The size of the second bar or pie of the bar-of-pie or pie-of-pie chart as a percentage of the size of the first pie. This ranges from 5% to 200% (both inclusive).
        /// </summary>
        public ushort SecondPieSize
        {
            get { return iSecondPieSize; }
            set
            {
                iSecondPieSize = value;
                if (iSecondPieSize < 5) iSecondPieSize = 5;
                if (iSecondPieSize > 200) iSecondPieSize = 200;
            }
        }

        internal SLA.SLShapeProperties ShapeProperties;

        /// <summary>
        /// Line properties for the connecting line for bar-of-pie or pie-of-pie charts.
        /// </summary>
        public SLA.SLLinePropertiesType Line { get { return this.ShapeProperties.Outline; } }

        /// <summary>
        /// Shadow properties for the connecting line for bar-of-pie or pie-of-pie charts.
        /// </summary>
        public SLA.SLShadowEffect Shadow { get { return this.ShapeProperties.EffectList.Shadow; } }

        /// <summary>
        /// Glow properties for the connecting line for bar-of-pie or pie-of-pie charts.
        /// </summary>
        public SLA.SLGlow Glow { get { return this.ShapeProperties.EffectList.Glow; } }

        /// <summary>
        /// Soft edge properties for the connecting line for bar-of-pie or pie-of-pie charts.
        /// </summary>
        public SLA.SLSoftEdge SoftEdge { get { return this.ShapeProperties.EffectList.SoftEdge; } }

        /// <summary>
        /// Initializes an instance of SLPieChartOptions. It is recommended to use SLChart.CreatePieChartOptions().
        /// </summary>
        public SLPieChartOptions()
        {
            this.Initialize(new List<System.Drawing.Color>(), false);
        }

        internal SLPieChartOptions(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            this.Initialize(ThemeColors, ThrowExceptionsIfAny);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            this.VaryColors = true;
            this.iFirstSliceAngle = 0;
            this.byHoleSize = 10;
            this.iGapWidth = 150;
            this.HasSplit = false;
            this.SplitType = C.SplitValues.Position;
            this.SplitPosition = 0;
            this.SecondPiePoints = new List<int>();
            this.iSecondPieSize = 75;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
        }

        /// <summary>
        /// Split the data series by position where the second plot contains the last N values. This is only for bar-of-pie or pie-of-pie charts.
        /// </summary>
        /// <param name="LastNValues">The last N values used in the second plot.</param>
        public void SplitSeriesByPosition(int LastNValues)
        {
            this.HasSplit = true;
            this.SplitType = C.SplitValues.Position;
            this.SplitPosition = (double)LastNValues;
            this.SecondPiePoints.Clear();
        }

        /// <summary>
        /// Split the data series by value where the second plot contains all values less than a maximum value. This is only for bar-of-pie or pie-of-pie charts.
        /// </summary>
        /// <param name="MaxValue">The maximum value.</param>
        public void SplitSeriesByValue(double MaxValue)
        {
            this.HasSplit = true;
            this.SplitType = C.SplitValues.Value;
            this.SplitPosition = MaxValue;
            this.SecondPiePoints.Clear();
        }

        /// <summary>
        /// Split the data series by percentage where the second plot contains all values less than a percentage of the sum. This is only for bar-of-pie or pie-of-pie charts.
        /// </summary>
        /// <param name="MaxPercentage">The maximum percentage of the sum.</param>
        public void SplitSeriesByPercentage(double MaxPercentage)
        {
            this.HasSplit = true;
            this.SplitType = C.SplitValues.Percent;
            this.SplitPosition = MaxPercentage;
            this.SecondPiePoints.Clear();
        }

        /// <summary>
        /// Split the data series by selecting data points for the second plot. This is only for bar-of-pie or pie-of-pie charts.
        /// </summary>
        /// <param name="DataPointIndices">The indices of the data points of the data series. The index is 1-based, so "1,3,4" sets the 1st, 3rd and 4th data point in the second plot.</param>
        public void SplitSeriesByCustom(params int[] DataPointIndices)
        {
            this.HasSplit = true;
            this.SplitType = C.SplitValues.Custom;
            this.SplitPosition = 0;
            this.SecondPiePoints.Clear();
            foreach (int i in DataPointIndices)
            {
                // indices should start from 1 onwards
                if (i > 0) this.SecondPiePoints.Add(i - 1);
            }
            this.SecondPiePoints.Sort();
        }

        internal SLPieChartOptions Clone()
        {
            SLPieChartOptions pco = new SLPieChartOptions(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
            pco.VaryColors = this.VaryColors;
            pco.iFirstSliceAngle = this.iFirstSliceAngle;
            pco.byHoleSize = this.byHoleSize;
            pco.iGapWidth = this.iGapWidth;
            pco.HasSplit = this.HasSplit;
            pco.SplitType = this.SplitType;
            pco.SplitPosition = this.SplitPosition;

            pco.SecondPiePoints = new List<int>();
            for (int i = 0; i < this.SecondPiePoints.Count; ++i)
            {
                pco.SecondPiePoints.Add(this.SecondPiePoints[i]);
            }

            pco.iSecondPieSize = this.iSecondPieSize;

            pco.ShapeProperties = this.ShapeProperties.Clone();

            return pco;
        }
    }
}
