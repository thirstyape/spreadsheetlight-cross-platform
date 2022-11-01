using System;
using System.Collections.Generic;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Chart customization options for stock charts.
    /// </summary>
    public class SLStockChartOptions
    {
        internal ushort iGapWidth;
        /// <summary>
        /// The gap width between columns as a percentage of column width, ranging between 0% and 500% (both inclusive). The default is 150%.
        /// This only applies when there's Volume data.
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

        internal sbyte byOverlap;
        /// <summary>
        /// The amount of overlapping for columns, ranging from -100 to 100 (both inclusive). The default is 0.
        /// This only applies when there's Volume data.
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

        internal SLA.SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Fill properties for Volume data.
        /// </summary>
        public SLA.SLFill Fill { get { return this.ShapeProperties.Fill; } }

        /// <summary>
        /// Border properties for Volume data.
        /// </summary>
        public SLA.SLLinePropertiesType Border { get { return this.ShapeProperties.Outline; } }

        /// <summary>
        /// Shadow properties for Volume data.
        /// </summary>
        public SLA.SLShadowEffect Shadow { get { return this.ShapeProperties.EffectList.Shadow; } }

        /// <summary>
        /// Glow properties for Volume data.
        /// </summary>
        public SLA.SLGlow Glow { get { return this.ShapeProperties.EffectList.Glow; } }

        /// <summary>
        /// Soft edge properties for Volume data.
        /// </summary>
        public SLA.SLSoftEdge SoftEdge { get { return this.ShapeProperties.EffectList.SoftEdge; } }

        /// <summary>
        /// 3D format properties for Volume data.
        /// </summary>
        public SLA.SLFormat3D Format3D { get { return this.ShapeProperties.Format3D; } }

        /// <summary>
        /// Indicates if the stock chart has drop lines.
        /// </summary>
        public bool HasDropLines { get; set; }

        /// <summary>
        /// Drop lines properties.
        /// </summary>
        public SLDropLines DropLines { get; set; }

        /// <summary>
        /// Indicates if the stock chart has high-low lines.
        /// </summary>
        public bool HasHighLowLines { get; set; }

        /// <summary>
        /// High-low lines properties.
        /// </summary>
        public SLHighLowLines HighLowLines { get; set; }

        /// <summary>
        /// Indicates if the stock chart has up-down bars.
        /// </summary>
        public bool HasUpDownBars { get; set; }

        /// <summary>
        /// Up-down bars properties.
        /// </summary>
        public SLUpDownBars UpDownBars { get; set; }

        /// <summary>
        /// Initializes an instance of SLStockChartOptions. It is recommended to use SLChart.CreateStockChartOptions().
        /// </summary>
        public SLStockChartOptions()
        {
            this.Initialize(new List<System.Drawing.Color>(), false, false);
        }

        internal SLStockChartOptions(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.Initialize(ThemeColors, IsStylish, ThrowExceptionsIfAny);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.iGapWidth = 150;
            this.byOverlap = 0;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
            this.HasDropLines = false;
            this.DropLines = new SLDropLines(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.HasHighLowLines = true;
            this.HighLowLines = new SLHighLowLines(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.HasUpDownBars = true;
            this.UpDownBars = new SLUpDownBars(ThemeColors, IsStylish, ThrowExceptionsIfAny);
        }

        internal SLStockChartOptions Clone()
        {
            SLStockChartOptions sco = new SLStockChartOptions();
            sco.iGapWidth = this.iGapWidth;
            sco.byOverlap = this.byOverlap;
            sco.ShapeProperties = this.ShapeProperties.Clone();
            sco.HasDropLines = this.HasDropLines;
            sco.DropLines = this.DropLines.Clone();
            sco.HasHighLowLines = this.HasHighLowLines;
            sco.HighLowLines = this.HighLowLines.Clone();
            sco.HasUpDownBars = this.HasUpDownBars;
            sco.UpDownBars = this.UpDownBars.Clone();

            return sco;
        }
    }
}
