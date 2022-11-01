using System;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Chart customization options for bubble charts.
    /// </summary>
    public class SLBubbleChartOptions
    {
        /// <summary>
        /// Specifies if the bubbles have a 3D effect.
        /// </summary>
        public bool Bubble3D { get; set; }

        internal uint iBubbleScale;
        /// <summary>
        /// Scale factor in percentage of the default size, ranging from 0% to 300% (both inclusive). The default is 100%.
        /// </summary>
        public uint BubbleScale
        {
            get { return iBubbleScale; }
            set
            {
                iBubbleScale = value;
                if (iBubbleScale > 300) iBubbleScale = 300;
            }
        }

        /// <summary>
        /// Specifies if negatively sized bubbles are shown.
        /// </summary>
        public bool ShowNegativeBubbles { get; set; }

        /// <summary>
        /// Specifies how bubble sizes relate to the presentation of the bubbles.
        /// </summary>
        public C.SizeRepresentsValues SizeRepresents { get; set; }

        /// <summary>
        /// Initializes an instance of SLBubbleChartOptions.
        /// </summary>
        public SLBubbleChartOptions()
        {
            this.Bubble3D = true;
            this.iBubbleScale = 100;
            this.ShowNegativeBubbles = true;
            this.SizeRepresents = C.SizeRepresentsValues.Area;
        }

        internal SLBubbleChartOptions Clone()
        {
            SLBubbleChartOptions bco = new SLBubbleChartOptions();
            bco.Bubble3D = this.Bubble3D;
            bco.iBubbleScale = this.iBubbleScale;
            bco.ShowNegativeBubbles = this.ShowNegativeBubbles;
            bco.SizeRepresents = this.SizeRepresents;

            return bco;
        }
    }
}
