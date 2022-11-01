using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting chart legends.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.Legend class.
    /// </summary>
    public class SLLegend
    {
        /// <summary>
        /// The position of the legend.
        /// </summary>
        public C.LegendPositionValues LegendPosition { get; set; }

        /// <summary>
        /// Specifies if the legend is overlayed. True if the legend overlaps the plot area, false otherwise.
        /// </summary>
        public bool Overlay { get; set; }

        internal SLA.SLShapeProperties ShapeProperties;

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

        internal SLLegend(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.LegendPosition = IsStylish ? C.LegendPositionValues.Bottom : C.LegendPositionValues.Right;
            this.Overlay = false;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);

            if (IsStylish)
            {
                this.ShapeProperties.Fill.SetNoFill();
                this.ShapeProperties.Outline.SetNoLine();
            }
            else
            {
                this.ShapeProperties.Fill.BlipDpi = 0;
                this.ShapeProperties.Fill.BlipRotateWithShape = true;
            }
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
        }

        internal C.Legend ToLegend(bool IsStylish)
        {
            C.Legend l = new C.Legend();
            l.LegendPosition = new C.LegendPosition() { Val = this.LegendPosition };

            l.Append(new C.Layout());
            l.Append(new C.Overlay() { Val = this.Overlay });

            if (this.ShapeProperties.HasShapeProperties) l.Append(this.ShapeProperties.ToChartShapeProperties(IsStylish));

            if (IsStylish)
            {
                C.TextProperties tp = new C.TextProperties();
                tp.BodyProperties = new A.BodyProperties()
                {
                    Rotation = 0,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                tp.ListStyle = new A.ListStyle();

                A.Paragraph para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();

                A.DefaultRunProperties defrunprops = new A.DefaultRunProperties();
                defrunprops.FontSize = 900;
                defrunprops.Bold = false;
                defrunprops.Italic = false;
                defrunprops.Underline = A.TextUnderlineValues.None;
                defrunprops.Strike = A.TextStrikeValues.NoStrike;
                defrunprops.Kerning = 1200;
                defrunprops.Baseline = 0;

                A.SchemeColor schclr = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
                schclr.Append(new A.LuminanceModulation() { Val = 65000 });
                schclr.Append(new A.LuminanceOffset() { Val = 35000 });
                defrunprops.Append(new A.SolidFill()
                {
                    SchemeColor = schclr
                });

                defrunprops.Append(new A.LatinFont() { Typeface = "+mn-lt" });
                defrunprops.Append(new A.EastAsianFont() { Typeface = "+mn-ea" });
                defrunprops.Append(new A.ComplexScriptFont() { Typeface = "+mn-cs" });

                para.ParagraphProperties.Append(defrunprops);
                para.Append(new A.EndParagraphRunProperties() { Language = System.Globalization.CultureInfo.CurrentCulture.Name });

                tp.Append(para);

                l.Append(tp);
            }

            return l;
        }

        internal SLLegend Clone()
        {
            SLLegend l = new SLLegend(this.ShapeProperties.listThemeColors, false, this.ShapeProperties.ThrowExceptionsIfAny);
            l.LegendPosition = this.LegendPosition;
            l.Overlay = this.Overlay;
            l.ShapeProperties = this.ShapeProperties.Clone();

            return l;
        }
    }
}
