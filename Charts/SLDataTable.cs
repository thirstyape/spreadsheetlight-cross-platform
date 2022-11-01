using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting the data table of charts.
    /// </summary>
    public class SLDataTable
    {
        internal SLA.SLShapeProperties ShapeProperties;

        /// <summary>
        /// Specifies if horizontal table borders are shown.
        /// </summary>
        public bool ShowHorizontalBorder { get; set; }

        /// <summary>
        /// Specifies if vertical table borders are shown.
        /// </summary>
        public bool ShowVerticalBorder { get; set; }

        /// <summary>
        /// Specifies if table outline borders are shown.
        /// </summary>
        public bool ShowOutlineBorder { get; set; }

        /// <summary>
        /// Specifies if legend keys are shown.
        /// </summary>
        public bool ShowLegendKeys { get; set; }

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

        internal SLFont Font { get; set; }

        internal SLDataTable(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.ShowHorizontalBorder = true;
            this.ShowVerticalBorder = true;
            this.ShowOutlineBorder = true;
            this.ShowLegendKeys = true;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);

            if (IsStylish)
            {
                this.ShapeProperties.Fill.SetNoFill();
                this.ShapeProperties.Outline.Width = 0.75m;
                this.ShapeProperties.Outline.CapType = A.LineCapValues.Flat;
                this.ShapeProperties.Outline.CompoundLineType = A.CompoundLineValues.Single;
                this.ShapeProperties.Outline.Alignment = A.PenAlignmentValues.Center;
                this.ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                this.ShapeProperties.Outline.JoinType = SLA.SLLineJoinValues.Round;
            }

            this.Font = null;
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
        }

        /// <summary>
        /// Set font settings for the contents of the data table.
        /// </summary>
        /// <param name="Font">The SLFont containing the font settings.</param>
        public void SetFont(SLFont Font)
        {
            this.Font = Font.Clone();
        }

        internal C.DataTable ToDataTable(bool IsStylish)
        {
            C.DataTable dt = new C.DataTable();

            if (this.ShowHorizontalBorder) dt.ShowHorizontalBorder = new C.ShowHorizontalBorder() { Val = true };
            if (this.ShowVerticalBorder) dt.ShowVerticalBorder = new C.ShowVerticalBorder() { Val = true };
            if (this.ShowOutlineBorder) dt.ShowOutlineBorder = new C.ShowOutlineBorder() { Val = true };
            if (this.ShowLegendKeys) dt.ShowKeys = new C.ShowKeys() { Val = true };

            if (this.ShapeProperties.HasShapeProperties) dt.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            if (this.Font != null)
            {
                dt.TextProperties = new C.TextProperties();
                dt.TextProperties.BodyProperties = new A.BodyProperties();
                dt.TextProperties.ListStyle = new A.ListStyle();

                dt.TextProperties.Append(this.Font.ToParagraph());
            }
            else if (IsStylish)
            {
                dt.TextProperties = new C.TextProperties();
                dt.TextProperties.BodyProperties = new A.BodyProperties()
                {
                    Rotation = 0,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                dt.TextProperties.ListStyle = new A.ListStyle();

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

                dt.TextProperties.Append(para);
            }

            return dt;
        }

        internal SLDataTable Clone()
        {
            SLDataTable dt = new SLDataTable(this.ShapeProperties.listThemeColors, false, this.ShapeProperties.ThrowExceptionsIfAny);
            dt.ShapeProperties = this.ShapeProperties.Clone();
            dt.ShowHorizontalBorder = this.ShowHorizontalBorder;
            dt.ShowVerticalBorder = this.ShowVerticalBorder;
            dt.ShowOutlineBorder = this.ShowOutlineBorder;
            dt.ShowLegendKeys = this.ShowLegendKeys;
            if (this.Font != null) dt.Font = this.Font.Clone();

            return dt;
        }
    }
}
