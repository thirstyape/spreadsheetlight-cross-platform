using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting data markers in charts.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.Marker class.
    /// </summary>
    public class SLMarker
    {
        internal bool HasMarker
        {
            get
            {
                return this.vSymbol != null || this.bySize != null || this.ShapeProperties.HasShapeProperties;
            }
        }

        internal C.MarkerStyleValues? vSymbol;
        /// <summary>
        /// Marker symbol.
        /// </summary>
        public C.MarkerStyleValues Symbol
        {
            get { return vSymbol ?? C.MarkerStyleValues.Auto; }
            set
            {
                vSymbol = value;
            }
        }

        internal byte? bySize;
        /// <summary>
        /// Range is 2 to 72 inclusive. Default is 5 in Open XML but Excel uses 7.
        /// </summary>
        public byte Size
        {
            get { return bySize ?? 5; }
            set
            {
                bySize = value;
                if (bySize != null)
                {
                    if (bySize < 2) bySize = 2;
                    if (bySize > 72) bySize = 72;
                }
            }
        }

        internal SLA.SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Fill properties.
        /// </summary>
        public SLA.SLFill Fill { get { return this.ShapeProperties.Fill; } }

        /// <summary>
        /// Line properties.
        /// </summary>
        public SLA.SLLinePropertiesType Line { get { return this.ShapeProperties.Outline; } }

        internal SLMarker(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            this.vSymbol = null;
            this.bySize = null;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
        }

        internal C.Marker ToMarker(bool IsStylish)
        {
            C.Marker m = new C.Marker();
            if (this.vSymbol != null) m.Symbol = new C.Symbol() { Val = this.vSymbol.Value };
            if (this.bySize != null) m.Size = new C.Size() { Val = this.bySize.Value };

            if (this.ShapeProperties.HasShapeProperties)
            {
                m.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            return m;
        }

        internal SLMarker Clone()
        {
            SLMarker m = new SLMarker(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
            m.Symbol = this.Symbol;
            m.bySize = this.bySize;
            m.ShapeProperties = this.ShapeProperties.Clone();

            return m;
        }
    }
}
