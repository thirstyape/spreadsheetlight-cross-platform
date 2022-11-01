using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// This simulates the element group EG_DLblShared as specified in the Open XML specs.
    /// </summary>
    public abstract class EGDLblShared : SLChartAlignment
    {
        // This is C.NumberingFormat
        internal bool HasNumberingFormat;

        internal string sFormatCode;
        /// <summary>
        /// Format code. If you set a custom format code, you might also want to set SourceLinked to false.
        /// </summary>
        public string FormatCode
        {
            get { return sFormatCode; }
            set
            {
                sFormatCode = value;
                HasNumberingFormat = true;
            }
        }

        internal bool bSourceLinked;
        /// <summary>
        /// Whether the format code is linked to the data source.
        /// </summary>
        public bool SourceLinked
        {
            get { return bSourceLinked; }
            set
            {
                bSourceLinked = value;
                HasNumberingFormat = true;
            }
        }

        internal C.DataLabelPositionValues? vLabelPosition;

        internal SLA.SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Fill properties.
        /// </summary>
        public SLA.SLFill Fill { get { return this.ShapeProperties.Fill; } }

        /// <summary>
        /// Border properties.
        /// </summary>
        public SLA.SLLinePropertiesType Border { get { return this.ShapeProperties.Outline; } }

        /// <summary>
        /// Shadow properties.
        /// </summary>
        public SLA.SLShadowEffect Shadow { get { return this.ShapeProperties.EffectList.Shadow; } }

        /// <summary>
        /// Glow properties.
        /// </summary>
        public SLA.SLGlow Glow { get { return this.ShapeProperties.EffectList.Glow; } }

        /// <summary>
        /// Soft edge properties.
        /// </summary>
        public SLA.SLSoftEdge SoftEdge { get { return this.ShapeProperties.EffectList.SoftEdge; } }

        /// <summary>
        /// 3D format properties.
        /// </summary>
        public SLA.SLFormat3D Format3D { get { return this.ShapeProperties.Format3D; } }

        /// <summary>
        /// Specifies if the legend key is included in the label.
        /// </summary>
        public bool ShowLegendKey { get; set; }

        /// <summary>
        /// Specifies if the label contains the value. For certain charts, this is known as the "Y Value".
        /// </summary>
        public bool ShowValue { get; set; }

        /// <summary>
        /// Specifies if the label contains the category name. For certain charts, this is known as the "X Value".
        /// </summary>
        public bool ShowCategoryName { get; set; }

        /// <summary>
        /// Specifies if the label contains the series name.
        /// </summary>
        public bool ShowSeriesName { get; set; }

        /// <summary>
        /// Specifies if the label contains the percentage. This is for pie charts.
        /// </summary>
        public bool ShowPercentage { get; set; }

        /// <summary>
        /// Specifies if the label contains the bubble size. This is for bubble charts.
        /// </summary>
        public bool ShowBubbleSize { get; set; }

        /// <summary>
        /// The separator.
        /// </summary>
        public string Separator { get; set; }

        internal EGDLblShared(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            this.sFormatCode = SLConstants.NumberFormatGeneral;
            this.bSourceLinked = true;
            this.HasNumberingFormat = false;
            this.vLabelPosition = null;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
            this.ShowLegendKey = false;
            this.ShowValue = false;
            this.ShowCategoryName = false;
            this.ShowSeriesName = false;
            this.ShowPercentage = false;
            this.ShowBubbleSize = false;
            this.Separator = string.Empty;
        }

        /// <summary>
        /// Set the position of the data label.
        /// </summary>
        /// <param name="Position">The data label position.</param>
        public void SetLabelPosition(C.DataLabelPositionValues Position)
        {
            this.vLabelPosition = Position;
        }

        /// <summary>
        /// Set automatic positioning of the data label.
        /// </summary>
        public void SetAutoLabelPosition()
        {
            this.vLabelPosition = null;
        }
    }
}
