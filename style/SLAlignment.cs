using System;
using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Specifies reading order.
    /// </summary>
    public enum SLAlignmentReadingOrderValues
    {
        /// <summary>
        /// Reading order is context dependent.
        /// </summary>
        ContextDependent = 0,
        /// <summary>
        /// Reading order is from left to right.
        /// </summary>
        LeftToRight = 1,
        /// <summary>
        /// Reading order is from right to left.
        /// </summary>
        RightToLeft = 2
    }

    /// <summary>
    /// Encapsulates properties and methods for text alignment in cells. This simulates the DocumentFormat.OpenXml.Spreadsheet.Alignment class.
    /// </summary>
    public class SLAlignment
    {
        internal bool HasHorizontal;
        private HorizontalAlignmentValues vHorizontal;
        /// <summary>
        /// Specifies the horizontal alignment. Default value is General.
        /// </summary>
        public HorizontalAlignmentValues Horizontal
        {
            get { return vHorizontal; }
            set
            {
                vHorizontal = value;
                HasHorizontal = true;
            }
        }

        internal bool HasVertical;
        private VerticalAlignmentValues vVertical;
        /// <summary>
        /// Specifies the vertical alignment. Default value is Bottom.
        /// </summary>
        public VerticalAlignmentValues Vertical
        {
            get { return vVertical; }
            set
            {
                vVertical = value;
                HasVertical = true;
            }
        }

        private int? iTextRotation;
        /// <summary>
        /// Specifies the rotation angle of the text, ranging from -90 degrees to 90 degrees. Default value is 0 degrees.
        /// </summary>
        public int? TextRotation
        {
            get { return iTextRotation; }
            set
            {
                if (value.HasValue && (value.Value >= -90) && (value.Value <= 90))
                {
                    iTextRotation = value.Value;
                }
                else
                {
                    iTextRotation = null;
                }
            }
        }

        /// <summary>
        /// Specifies if the text in the cell should be wrapped.
        /// </summary>
        public bool? WrapText { get; set; }

        /// <summary>
        /// Specifies the indent. Each unit value equals 3 spaces.
        /// </summary>
        public uint? Indent { get; set; }

        /// <summary>
        /// This property is used when the class is part of a SLDifferentialFormat class. It specifies the indent value in addition to the given Indent property.
        /// </summary>
        public int? RelativeIndent { get; set; }

        /// <summary>
        /// Specifies if the last line should be justified (usually for East Asian fonts).
        /// </summary>
        public bool? JustifyLastLine { get; set; }

        /// <summary>
        /// Specifies if the text in the cell should be shrunk to fit the cell.
        /// </summary>
        public bool? ShrinkToFit { get; set; }

        internal bool HasReadingOrder;
        private SLAlignmentReadingOrderValues vReadingOrder;
        /// <summary>
        /// Specifies the reading order of the text in the cell.
        /// </summary>
        public SLAlignmentReadingOrderValues ReadingOrder
        {
            get { return vReadingOrder; }
            set
            {
                vReadingOrder = value;
                HasReadingOrder = true;
            }
        }

        /// <summary>
        /// Initializes an instance of SLAlignment.
        /// </summary>
        public SLAlignment()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.vHorizontal = HorizontalAlignmentValues.General;
            this.HasHorizontal = false;
            this.vVertical = VerticalAlignmentValues.Bottom;
            this.HasVertical = false;
            this.TextRotation = null;
            this.WrapText = null;
            this.Indent = null;
            this.RelativeIndent = null;
            this.JustifyLastLine = null;
            this.ShrinkToFit = null;
            this.vReadingOrder = SLAlignmentReadingOrderValues.LeftToRight;
            this.HasReadingOrder = false;
        }

        internal void FromAlignment(Alignment align)
        {
            this.SetAllNull();

            if (align.Horizontal != null) this.Horizontal = align.Horizontal.Value;
            if (align.Vertical != null) this.Vertical = align.Vertical.Value;

            if (align.TextRotation != null && align.TextRotation.Value <= 180)
            {
                this.TextRotation = this.TextRotationToIntuitiveValue(align.TextRotation.Value);
            }

            if (align.WrapText != null) this.WrapText = align.WrapText.Value;
            if (align.Indent != null) this.Indent = align.Indent.Value;
            if (align.RelativeIndent != null) this.RelativeIndent = align.RelativeIndent.Value;
            if (align.JustifyLastLine != null) this.JustifyLastLine = align.JustifyLastLine.Value;
            if (align.ShrinkToFit != null) this.ShrinkToFit = align.ShrinkToFit.Value;

            if (align.ReadingOrder != null)
            {
                switch (align.ReadingOrder.Value)
                {
                    case (uint)SLAlignmentReadingOrderValues.ContextDependent:
                        this.ReadingOrder = SLAlignmentReadingOrderValues.ContextDependent;
                        break;
                    case (uint)SLAlignmentReadingOrderValues.LeftToRight:
                        this.ReadingOrder = SLAlignmentReadingOrderValues.LeftToRight;
                        break;
                    case (uint)SLAlignmentReadingOrderValues.RightToLeft:
                        this.ReadingOrder = SLAlignmentReadingOrderValues.RightToLeft;
                        break;
                }
            }
        }

        internal Alignment ToAlignment()
        {
            Alignment align = new Alignment();
            if (this.HasHorizontal) align.Horizontal = this.Horizontal;
            if (this.HasVertical) align.Vertical = this.Vertical;
            if (this.TextRotation != null) align.TextRotation = this.TextRotationToOpenXmlValue(this.TextRotation.Value);
            if (this.WrapText != null) align.WrapText = this.WrapText.Value;
            if (this.Indent != null) align.Indent = this.Indent.Value;
            if (this.RelativeIndent != null) align.RelativeIndent = this.RelativeIndent.Value;
            if (this.JustifyLastLine != null) align.JustifyLastLine = this.JustifyLastLine.Value;
            if (this.ShrinkToFit != null) align.ShrinkToFit = this.ShrinkToFit.Value;
            if (this.HasReadingOrder) align.ReadingOrder = (uint)this.ReadingOrder;

            return align;
        }

        internal void FromHash(string Hash)
        {
            this.SetAllNull();

            string[] sa = Hash.Split(new string[] { SLConstants.XmlAlignmentAttributeSeparator }, StringSplitOptions.None);

            if (sa.Length >= 9)
            {
                if (!sa[0].Equals("null")) this.Horizontal = (HorizontalAlignmentValues)Enum.Parse(typeof(HorizontalAlignmentValues), sa[0]);

                if (!sa[1].Equals("null")) this.Vertical = (VerticalAlignmentValues)Enum.Parse(typeof(VerticalAlignmentValues), sa[1]);

                if (!sa[2].Equals("null")) this.TextRotation = int.Parse(sa[2]);

                if (!sa[3].Equals("null")) this.WrapText = bool.Parse(sa[3]);

                if (!sa[4].Equals("null")) this.Indent = uint.Parse(sa[4]);

                if (!sa[5].Equals("null")) this.RelativeIndent = int.Parse(sa[5]);

                if (!sa[6].Equals("null")) this.JustifyLastLine = bool.Parse(sa[6]);

                if (!sa[7].Equals("null")) this.ShrinkToFit = bool.Parse(sa[7]);

                if (!sa[8].Equals("null")) this.ReadingOrder = (SLAlignmentReadingOrderValues)Enum.Parse(typeof(SLAlignmentReadingOrderValues), sa[8]);
            }
        }

        internal string ToHash()
        {
            StringBuilder sb = new StringBuilder();

            if (this.HasHorizontal) sb.AppendFormat("{0}{1}", this.Horizontal, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (this.HasVertical) sb.AppendFormat("{0}{1}", this.Vertical, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (this.TextRotation != null) sb.AppendFormat("{0}{1}", this.TextRotation.Value.ToString(CultureInfo.InvariantCulture), SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (this.WrapText != null) sb.AppendFormat("{0}{1}", this.WrapText.Value, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (this.Indent != null) sb.AppendFormat("{0}{1}", this.Indent.Value.ToString(CultureInfo.InvariantCulture), SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (this.RelativeIndent != null) sb.AppendFormat("{0}{1}", this.RelativeIndent.Value.ToString(CultureInfo.InvariantCulture), SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (this.JustifyLastLine != null) sb.AppendFormat("{0}{1}", this.JustifyLastLine.Value, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (this.ShrinkToFit != null) sb.AppendFormat("{0}{1}", this.ShrinkToFit.Value, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (this.HasReadingOrder) sb.AppendFormat("{0}{1}", this.ReadingOrder, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            return sb.ToString();
        }

        internal int TextRotationToIntuitiveValue(uint Degree)
        {
            int iDegree = 0;

            if (Degree >= 0 && Degree <= 90)
            {
                iDegree = (int)Degree;
            }
            else if (Degree >= 91 && Degree <= 180)
            {
                iDegree = 90 - (int)Degree;
            }

            return iDegree;
        }

        internal uint TextRotationToOpenXmlValue(int Degree)
        {
            uint iDegree = 0;

            if (Degree >= 0 && Degree <= 90)
            {
                iDegree = (uint)Degree;
            }
            else if (Degree >= -90 && Degree < 0)
            {
                iDegree = (uint)(90 - Degree);
            }

            return iDegree;
        }

        internal string WriteToXmlTag()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<x:alignment");

            if (this.HasHorizontal)
            {
                switch (this.Horizontal)
                {
                    case HorizontalAlignmentValues.Center:
                        sb.Append(" horizontal=\"center\"");
                        break;
                    case HorizontalAlignmentValues.CenterContinuous:
                        sb.Append(" horizontal=\"centerContinuous\"");
                        break;
                    case HorizontalAlignmentValues.Distributed:
                        sb.Append(" horizontal=\"distributed\"");
                        break;
                    case HorizontalAlignmentValues.Fill:
                        sb.Append(" horizontal=\"fill\"");
                        break;
                    case HorizontalAlignmentValues.General:
                        sb.Append(" horizontal=\"general\"");
                        break;
                    case HorizontalAlignmentValues.Justify:
                        sb.Append(" horizontal=\"justify\"");
                        break;
                    case HorizontalAlignmentValues.Left:
                        sb.Append(" horizontal=\"left\"");
                        break;
                    case HorizontalAlignmentValues.Right:
                        sb.Append(" horizontal=\"right\"");
                        break;
                }
            }

            if (this.HasVertical)
            {
                switch (this.Vertical)
                {
                    case VerticalAlignmentValues.Bottom:
                        sb.Append(" vertical=\"bottom\"");
                        break;
                    case VerticalAlignmentValues.Center:
                        sb.Append(" vertical=\"center\"");
                        break;
                    case VerticalAlignmentValues.Distributed:
                        sb.Append(" vertical=\"distributed\"");
                        break;
                    case VerticalAlignmentValues.Justify:
                        sb.Append(" vertical=\"justify\"");
                        break;
                    case VerticalAlignmentValues.Top:
                        sb.Append(" vertical=\"top\"");
                        break;
                }
            }

            if (this.TextRotation != null) sb.AppendFormat(" textRotation=\"{0}\"", this.TextRotationToOpenXmlValue(this.TextRotation.Value).ToString(CultureInfo.InvariantCulture));
            if (this.WrapText != null) sb.AppendFormat(" wrapText=\"{0}\"", this.WrapText.Value ? "1" : "0");
            if (this.Indent != null) sb.AppendFormat(" indent=\"{0}\"", this.Indent.Value.ToString(CultureInfo.InvariantCulture));
            if (this.RelativeIndent != null) sb.AppendFormat(" relativeIndent=\"{0}\"", this.RelativeIndent.Value.ToString(CultureInfo.InvariantCulture));
            if (this.JustifyLastLine != null) sb.AppendFormat(" justifyLastLine=\"{0}\"", this.JustifyLastLine.Value ? "1" : "0");
            if (this.ShrinkToFit != null) sb.AppendFormat(" shrinkToFit=\"{0}\"", this.ShrinkToFit.Value ? "1" : "0");
            if (this.HasReadingOrder) sb.AppendFormat(" readingOrder=\"{0}\"", (uint)this.ReadingOrder);

            sb.Append(" />");

            return sb.ToString();
        }

        internal SLAlignment Clone()
        {
            SLAlignment align = new SLAlignment();
            align.HasHorizontal = this.HasHorizontal;
            align.vHorizontal = this.vHorizontal;
            align.HasVertical = this.HasVertical;
            align.vVertical = this.vVertical;
            align.iTextRotation = this.iTextRotation;
            align.WrapText = this.WrapText;
            align.Indent = this.Indent;
            align.RelativeIndent = this.RelativeIndent;
            align.JustifyLastLine = this.JustifyLastLine;
            align.ShrinkToFit = this.ShrinkToFit;
            align.HasReadingOrder = this.HasReadingOrder;
            align.vReadingOrder = this.vReadingOrder;

            return align;
        }
    }
}
