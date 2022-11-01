using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Vml;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight
{
    /// <summary>
    /// Specifies how the text is aligned horizontally.
    /// </summary>
    public enum SLHorizontalTextAlignmentValues
    {
        /// <summary>
        /// Left
        /// </summary>
        Left = 0,
        /// <summary>
        /// Justify
        /// </summary>
        Justify,
        /// <summary>
        /// Center
        /// </summary>
        Center,
        /// <summary>
        /// Right
        /// </summary>
        Right,
        /// <summary>
        /// Distributed
        /// </summary>
        Distributed
    }

    /// <summary>
    /// Specifies how the text is aligned vertically.
    /// </summary>
    public enum SLVerticalTextAlignmentValues
    {
        /// <summary>
        /// Top
        /// </summary>
        Top = 0,
        /// <summary>
        /// Justify
        /// </summary>
        Justify,
        /// <summary>
        /// Center
        /// </summary>
        Center,
        /// <summary>
        /// Bottom
        /// </summary>
        Bottom,
        /// <summary>
        /// Distributed
        /// </summary>
        Distributed
    }

    /// <summary>
    /// Specifies how the comment is oriented.
    /// </summary>
    public enum SLCommentOrientationValues
    {
        /// <summary>
        /// Horizontal
        /// </summary>
        Horizontal = 0,
        /// <summary>
        /// The text characters are arranged in a top-down direction
        /// </summary>
        TopDown,
        /// <summary>
        /// Rotated 270 degrees
        /// </summary>
        Rotated270Degrees,
        /// <summary>
        /// Rotated 90 degrees
        /// </summary>
        Rotated90Degrees
    }

    /// <summary>
    /// Specifies how line dashes are styled
    /// </summary>
    public enum SLDashStyleValues
    {
        /// <summary>
        /// Solid
        /// </summary>
        Solid = 0,
        /// <summary>
        /// Short dash
        /// </summary>
        ShortDash,
        /// <summary>
        /// Short dot
        /// </summary>
        ShortDot,
        /// <summary>
        /// Short dash dot
        /// </summary>
        ShortDashDot,
        /// <summary>
        /// Short dash dot dot
        /// </summary>
        ShortDashDotDot,
        /// <summary>
        /// Dot
        /// </summary>
        Dot,
        /// <summary>
        /// Dash
        /// </summary>
        Dash,
        /// <summary>
        /// Long dash
        /// </summary>
        LongDash,
        /// <summary>
        /// Dash dot
        /// </summary>
        DashDot,
        /// <summary>
        /// Long dash dot
        /// </summary>
        LongDashDot,
        /// <summary>
        /// Long dash dot dot
        /// </summary>
        LongDashDotDot
    }

    /// <summary>
    /// Encapsulates properties and methods for cell comments.
    /// </summary>
    public class SLComment
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal bool ThrowExceptionsIfAny { get; set; }

        // TODO: move with cells and size with cells

        internal string sAuthor;
        /// <summary>
        /// The author of the comment.
        /// </summary>
        public string Author
        {
            get { return sAuthor; }
            set { sAuthor = value.Trim(); }
        }

        internal SLRstType rst;

        internal bool HasSetPosition;

        internal double Top { get; set; }
        internal double Left { get; set; }

        internal bool UsePositionMargin;
        internal double TopMargin { get; set; }
        internal double LeftMargin { get; set; }

        /// <summary>
        /// Set true to automatically size the comment box according to the comment's contents.
        /// </summary>
        public bool AutoSize { get; set; }

        internal double fWidth;
        /// <summary>
        /// Width of comment box in units of points. For practical purposes, the width is a minimum of 1 pt.
        /// </summary>
        public double Width
        {
            get { return fWidth; }
            set
            {
                fWidth = value;
                if (fWidth < 1.0) fWidth = 1.0;
                AutoSize = false;
            }
        }

        internal double fHeight;
        /// <summary>
        /// Height of comment box in units of points. For practical purposes, the height is a minimum of 1 pt.
        /// </summary>
        public double Height
        {
            get { return fHeight; }
            set
            {
                fHeight = value;
                if (fHeight < 1.0) fHeight = 1.0;
                AutoSize = false;
            }
        }

        /// <summary>
        /// Fill properties. Note that this is repurposed, and some of the methods and properties can't be
        /// directly translated to a VML-equivalent (which is how comment styles are stored).
        /// </summary>
        public SLA.SLFill Fill { get; set; }

        internal byte bFromTransparency;
        internal byte bToTransparency;

        /// <summary>
        /// The transparency value of the first gradient point measured in percentage, ranging from 0% to 100% (both inclusive).
        /// </summary>
        public byte GradientFromTransparency
        {
            get { return bFromTransparency; }
            set
            {
                bFromTransparency = value;
                if (bFromTransparency > 100) bFromTransparency = 100;
            }
        }

        /// <summary>
        /// The transparency value of the last gradient point measured in percentage, ranging from 0% to 100% (both inclusive).
        /// </summary>
        public byte GradientToTransparency
        {
            get { return bToTransparency; }
            set
            {
                bToTransparency = value;
                if (bToTransparency > 100) bToTransparency = 100;
            }
        }

        /// <summary>
        /// Set null for automatic color.
        /// </summary>
        public System.Drawing.Color? LineColor { get; set; }

        internal double? fLineWeight;
        /// <summary>
        /// Line weight in points.
        /// </summary>
        public double LineWeight
        {
            // 0.75pt seems to be Excel's default, although the Open XML specs state 1pt as the default
            get { return fLineWeight ?? 0.75; }
            set
            {
                fLineWeight = value;
                if (fLineWeight < 0) fLineWeight = 0;
            }
        }

        /// <summary>
        /// Line style.
        /// </summary>
        public StrokeLineStyleValues LineStyle { get; set; }

        internal SLDashStyleValues? vLineDashStyle;
        internal StrokeEndCapValues? vEndCap;

        /// <summary>
        /// Horizontal text alignment.
        /// </summary>
        public SLHorizontalTextAlignmentValues HorizontalTextAlignment { get; set; }

        /// <summary>
        /// Vertical text alignment.
        /// </summary>
        public SLVerticalTextAlignmentValues VerticalTextAlignment { get; set; }

        /// <summary>
        /// Comment text orientation.
        /// </summary>
        public SLCommentOrientationValues Orientation { get; set; }

        /// <summary>
        /// Comment text direction.
        /// </summary>
        public SLAlignmentReadingOrderValues TextDirection { get; set; }

        /// <summary>
        /// Specifies whether the comment box has a shadow.
        /// </summary>
        public bool HasShadow { get; set; }

        /// <summary>
        /// Specifies the color of the comment box's shadow.
        /// </summary>
        public System.Drawing.Color ShadowColor { get; set; }

        /// <summary>
        /// Specifies whether the comment is visible.
        /// </summary>
        public bool Visible { get; set; }

        internal SLComment(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }

            this.ThrowExceptionsIfAny = ThrowExceptionsIfAny;

            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.sAuthor = string.Empty;
            this.rst = new SLRstType();
            this.HasSetPosition = false;
            this.Top = 0;
            this.Left = 0;
            this.UsePositionMargin = false;
            this.TopMargin = 0;
            this.LeftMargin = 0;
            this.AutoSize = false;
            this.fWidth = SLConstants.DefaultCommentBoxWidth;
            this.fHeight = SLConstants.DefaultCommentBoxHeight;

            this.Fill = new SLA.SLFill(this.listThemeColors, this.ThrowExceptionsIfAny);
            this.Fill.SetSolidFill(System.Drawing.Color.FromArgb(255, 255, 225), 0);
            this.bFromTransparency = 0;
            this.bToTransparency = 0;

            this.LineColor = null;
            this.fLineWeight = null;
            this.LineStyle = StrokeLineStyleValues.Single;
            this.vLineDashStyle = null;
            this.vEndCap = null;
            this.HorizontalTextAlignment = SLHorizontalTextAlignmentValues.Left;
            this.VerticalTextAlignment = SLVerticalTextAlignmentValues.Top;
            this.Orientation = SLCommentOrientationValues.Horizontal;
            this.TextDirection = SLAlignmentReadingOrderValues.ContextDependent;

            this.HasShadow = true;
            this.ShadowColor = System.Drawing.Color.Black;

            this.Visible = false;
        }

        /// <summary>
        /// Set the comment text.
        /// </summary>
        /// <param name="Text">The comment text.</param>
        public void SetText(string Text)
        {
            this.rst = new SLRstType();
            this.rst.SetText(Text);
        }

        /// <summary>
        /// Set the comment text given rich text content.
        /// </summary>
        /// <param name="RichText">The rich text content</param>
        public void SetText(SLRstType RichText)
        {
            this.rst = new SLRstType();
            this.rst = RichText.Clone();
        }
        
        /// <summary>
        /// Set the position of the comment box. NOTE: This isn't an exact science. The positioning depends on the DPI of the computer's screen.
        /// </summary>
        /// <param name="Top">Top position of the comment box based on row index. For example, 0.5 means at the half-way point of the 1st row, 2.5 means at the half-way point of the 3rd row.</param>
        /// <param name="Left">Left position of the comment box based on column index. For example, 0.5 means at the half-way point of the 1st column, 2.5 means at the half-way point of the 3rd column.</param>
        public void SetPosition(double Top, double Left)
        {
            this.HasSetPosition = true;
            this.Top = Top;
            this.Left = Left;
        }

        /// <summary>
        /// Set the position of the comment box given the top and left margins measured in points. It is suggested to use SetPosition() instead. This method is provided as a means of convenience. NOTE: This isn't an exact science. The positioning depends on the DPI of the computer's screen.
        /// </summary>
        /// <param name="TopMargin">Top margin in points. This is measured from the top-left corner of the cell A1.</param>
        /// <param name="LeftMargin">Left margin in points. This is measured from the top-left corner of the cell A1.</param>
        public void SetPositionMargin(double TopMargin, double LeftMargin)
        {
            this.HasSetPosition = true;
            this.UsePositionMargin = true;
            this.TopMargin = TopMargin;
            this.LeftMargin = LeftMargin;
        }

        /// <summary>
        /// Set the dash style of the comment box.
        /// </summary>
        /// <param name="DashStyle">The dash style.</param>
        public void SetDashStyle(SLDashStyleValues DashStyle)
        {
            this.vLineDashStyle = DashStyle;
            this.vEndCap = null;
        }

        /// <summary>
        /// Set the dash style of the comment box.
        /// </summary>
        /// <param name="DashStyle">The dash style.</param>
        /// <param name="EndCap">The end cap of the lines.</param>
        public void SetDashStyle(SLDashStyleValues DashStyle, StrokeEndCapValues EndCap)
        {
            this.vLineDashStyle = DashStyle;
            this.vEndCap = EndCap;
        }

        internal SLComment Clone()
        {
            SLComment comm = new SLComment(this.listThemeColors, this.ThrowExceptionsIfAny);
            comm.sAuthor = this.sAuthor;
            comm.rst = this.rst.Clone();
            comm.HasSetPosition = this.HasSetPosition;
            comm.Top = this.Top;
            comm.Left = this.Left;
            comm.UsePositionMargin = this.UsePositionMargin;
            comm.TopMargin = this.TopMargin;
            comm.LeftMargin = this.LeftMargin;
            comm.AutoSize = this.AutoSize;
            comm.fWidth = this.fWidth;
            comm.fHeight = this.fHeight;
            comm.Fill = this.Fill.Clone();
            comm.bFromTransparency = this.bFromTransparency;
            comm.bToTransparency = this.bToTransparency;
            comm.LineColor = this.LineColor;
            comm.fLineWeight = this.fLineWeight;
            comm.LineStyle = this.LineStyle;
            comm.vLineDashStyle = this.vLineDashStyle;
            comm.vEndCap = this.vEndCap;
            comm.HorizontalTextAlignment = this.HorizontalTextAlignment;
            comm.VerticalTextAlignment = this.VerticalTextAlignment;
            comm.Orientation = this.Orientation;
            comm.TextDirection = this.TextDirection;
            comm.HasShadow = this.HasShadow;
            comm.ShadowColor = this.ShadowColor;
            comm.Visible = this.Visible;

            return comm;
        }
    }
}
