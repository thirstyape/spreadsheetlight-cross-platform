using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    public class SLTrendline
    {
        /// <summary>
        /// Trendline type.
        /// </summary>
        public C.TrendlineValues TrendlineType { get; set; }

        internal byte byPolynomialOrder = 0;
        /// <summary>
        /// Polynomial trendline order. This is the power of the polynomial, and is between 2 (default value) and 6 (both inclusive). This is only applicable if the trendline type is polynomial.
        /// </summary>
        public byte PolynomialOrder
        {
            get
            {
                return byPolynomialOrder;
            }
            set
            {
                if (value < 2) byPolynomialOrder = 2;
                else if (value > 6) byPolynomialOrder = 6;
                else byPolynomialOrder = value;
            }
        }

        internal uint iPeriod = 2;
        /// <summary>
        /// Moving average period, and is only 2 (default value) or 3. THis is only applicable if the trendline type is "moving average".
        /// </summary>
        public uint Period
        {
            get
            {
                return iPeriod;
            }
            set
            {
                if (value < 2) iPeriod = 2;
                else if (value > 3) iPeriod = 3;
                else iPeriod = value;
            }
        }

        internal string AutomaticTrendlineName { get; set; }
        // has a PivotSource sub-property, but we ignore for now till we support pivot tables...
        /// <summary>
        /// Trendline name.
        /// </summary>
        public string TrendlineName { get; set; }

        /// <summary>
        /// Forward period. Default value is 0 (if null it's also 0). From Open XML SDK documentation: This is the number of categories (or units on a scatter chart) that the trendline extends after the data for the series that is being trended. On non-scatter charts, the value shall be a multiple of 0.5. (Excel seems to NOT adhere to the restriction though...)
        /// </summary>
        public bool? Forward { get; set; }

        /// <summary>
        /// Backward period. Default value is 0 (if null it's also 0). From Open XML SDK documentation: This is the number of categories (or units on a scatter chart) that the trend line extends before the data for the series that is being trended. On non-scatter charts, the value shall be 0 or 0.5. (Excel seems to NOT adhere to the restriction though...)
        /// </summary>
        public double? Backward { get; set; }

        /// <summary>
        /// The intercept value where the trendline crosses the Y-axis. This is only supported if the trendline type is exponential, linear or polynomial. Default value is 0 (if null then no intercept value is set).
        /// </summary>
        public bool? Intercept { get; set; }

        /// <summary>
        /// Set true to display equation on chart. False otherwise.
        /// </summary>
        public bool DisplayEquation { get; set; }

        /// <summary>
        /// Set true to display R-squared value on chart. False otherwise.
        /// </summary>
        public bool DisplayRSquared { get; set; }

        internal SLTrendlineLabel TrendlineLabel { get; set; }

        /// <summary>
        /// Fill properties.
        /// </summary>
        public SLA.SLFill Fill { get { return this.TrendlineLabel.Fill; } }

        /// <summary>
        /// Border properties.
        /// </summary>
        public SLA.SLLinePropertiesType Border { get { return this.TrendlineLabel.Border; } }

        /// <summary>
        /// Shadow properties.
        /// </summary>
        public SLA.SLShadowEffect Shadow { get { return this.TrendlineLabel.Shadow; } }

        /// <summary>
        /// Glow properties.
        /// </summary>
        public SLA.SLGlow Glow { get { return this.TrendlineLabel.Glow; } }

        /// <summary>
        /// Soft edge properties.
        /// </summary>
        public SLA.SLSoftEdge SoftEdge { get { return this.TrendlineLabel.SoftEdge; } }

        /// <summary>
        /// 3D format properties.
        /// </summary>
        public SLA.SLFormat3D Format3D { get { return this.TrendlineLabel.Format3D; } }

        public SLTrendline()
        {
        }

        internal C.Trendline ToTrendline(bool IsStylish)
        {
            C.Trendline tl = new C.Trendline();
            tl.DisplayEquation = new C.DisplayEquation() { Val = true };
            tl.TrendlineType = new C.TrendlineType() { Val = C.TrendlineValues.Linear };

            tl.TrendlineLabel = this.TrendlineLabel.ToTrendlineLabel(IsStylish);

            return tl;
        }
    }
}
