using System;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting alignment in charts.
    /// </summary>
    public abstract class SLChartAlignment
    {
        internal decimal? Rotation { get; set; }
        internal A.TextVerticalValues? Vertical { get; set; }
        internal A.TextAnchoringTypeValues? Anchor { get; set; }
        internal bool? AnchorCenter { get; set; }

        /// <summary>
        /// Initializes an instance of SLChartAlignment.
        /// </summary>
        public SLChartAlignment()
        {
            this.RemoveTextAlignment();
        }

        /// <summary>
        /// Set a horizontal text direction.
        /// </summary>
        /// <param name="TextAlignment">The vertical text alignment in horizontal direction.</param>
        /// <param name="CustomAngle">Rotation angle, ranging from -90 to 90 degrees. Accurate to 1/60000 of a degree.</param>
        public void SetHorizontalTextDirection(SLA.SLTextVerticalAlignment TextAlignment, decimal CustomAngle)
        {
            if (CustomAngle < -90m) CustomAngle = -90m;
            if (CustomAngle > 90m) CustomAngle = 90m;

            // vertical axis having 0 degrees won't have the text horizontal.
            // So don't set null?
            //if (CustomAngle == 0m) this.Rotation = null;
            //else this.Rotation = CustomAngle;

            //if (CustomAngle == 0m) this.Vertical = null;
            //else this.Vertical = A.TextVerticalValues.Horizontal;

            this.Rotation = CustomAngle;
            this.Vertical = A.TextVerticalValues.Horizontal;

            switch (TextAlignment)
            {
                case SLA.SLTextVerticalAlignment.Top:
                    this.Anchor = A.TextAnchoringTypeValues.Top;
                    this.AnchorCenter = false;
                    break;
                case SLA.SLTextVerticalAlignment.Middle:
                    this.Anchor = A.TextAnchoringTypeValues.Center;
                    this.AnchorCenter = false;
                    break;
                case SLA.SLTextVerticalAlignment.Bottom:
                    this.Anchor = A.TextAnchoringTypeValues.Bottom;
                    this.AnchorCenter = false;
                    break;
                case SLA.SLTextVerticalAlignment.TopCentered:
                    this.Anchor = A.TextAnchoringTypeValues.Top;
                    this.AnchorCenter = true;
                    break;
                case SLA.SLTextVerticalAlignment.MiddleCentered:
                    this.Anchor = A.TextAnchoringTypeValues.Center;
                    this.AnchorCenter = true;
                    break;
                case SLA.SLTextVerticalAlignment.BottomCentered:
                    this.Anchor = A.TextAnchoringTypeValues.Bottom;
                    this.AnchorCenter = true;
                    break;
            }
        }

        /// <summary>
        /// Set a stacked (vertical) text direction.
        /// </summary>
        /// <param name="TextAlignment">The horizontal text alignment in vertical direction.</param>
        /// <param name="LeftToRight">True if the text runs left-to-right. False if the text runs right-to-left.</param>
        public void SetStackedTextDirection(SLA.SLTextHorizontalAlignment TextAlignment, bool LeftToRight)
        {
            this.Rotation = 0m;

            this.Vertical = LeftToRight ? A.TextVerticalValues.WordArtVertical : A.TextVerticalValues.WordArtLeftToRight;

            switch (TextAlignment)
            {
                case SLA.SLTextHorizontalAlignment.Left:
                    if (LeftToRight)
                    {
                        this.Anchor = A.TextAnchoringTypeValues.Top;
                        this.AnchorCenter = false;
                    }
                    else
                    {
                        this.Anchor = A.TextAnchoringTypeValues.Bottom;
                        this.AnchorCenter = false;
                    }
                    break;
                case SLA.SLTextHorizontalAlignment.Center:
                    this.Anchor = A.TextAnchoringTypeValues.Center;
                    this.AnchorCenter = false;
                    break;
                case SLA.SLTextHorizontalAlignment.Right:
                    if (LeftToRight)
                    {
                        this.Anchor = A.TextAnchoringTypeValues.Bottom;
                        this.AnchorCenter = false;
                    }
                    else
                    {
                        this.Anchor = A.TextAnchoringTypeValues.Top;
                        this.AnchorCenter = false;
                    }
                    break;
                case SLA.SLTextHorizontalAlignment.LeftMiddle:
                    if (LeftToRight)
                    {
                        this.Anchor = A.TextAnchoringTypeValues.Top;
                        this.AnchorCenter = false;
                    }
                    else
                    {
                        this.Anchor = A.TextAnchoringTypeValues.Bottom;
                        this.AnchorCenter = false;
                    }
                    break;
                case SLA.SLTextHorizontalAlignment.CenterMiddle:
                    this.Anchor = A.TextAnchoringTypeValues.Center;
                    this.AnchorCenter = true;
                    break;
                case SLA.SLTextHorizontalAlignment.RightMiddle:
                    if (LeftToRight)
                    {
                        this.Anchor = A.TextAnchoringTypeValues.Bottom;
                        this.AnchorCenter = true;
                    }
                    else
                    {
                        this.Anchor = A.TextAnchoringTypeValues.Top;
                        this.AnchorCenter = true;
                    }
                    break;
            }
        }

        /// <summary>
        /// Set the text rotated 90 degrees.
        /// </summary>
        public void SetTextRotated90Degrees()
        {
            this.Rotation = 90m;
            this.Vertical = A.TextVerticalValues.Horizontal;
            this.Anchor = A.TextAnchoringTypeValues.Top;
            this.AnchorCenter = false;
        }

        /// <summary>
        /// Set the text rotated 270 degrees.
        /// </summary>
        public void SetTextRotated270Degrees()
        {
            this.Rotation = -90m;
            this.Vertical = A.TextVerticalValues.Horizontal;
            this.Anchor = A.TextAnchoringTypeValues.Top;
            this.AnchorCenter = false;
        }

        /// <summary>
        /// Remove all text alignment.
        /// </summary>
        public void RemoveTextAlignment()
        {
            this.Rotation = null;
            this.Vertical = null;
            this.Anchor = null;
            this.AnchorCenter = null;
        }
    }
}
