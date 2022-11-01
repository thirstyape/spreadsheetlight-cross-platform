using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for specifying incremental formatting. This simulates the DocumentFormat.OpenXml.Spreadsheet.DifferentialFormat and DocumentFormat.OpenXml.Office2010.Excel.DifferentialType classes.
    /// </summary>
    public class SLDifferentialFormat
    {
        internal bool HasAlignment;
        private SLAlignment alignReal;
        /// <summary>
        /// The alignment for incremental formatting.
        /// </summary>
        public SLAlignment Alignment
        {
            get { return alignReal; }
            set
            {
                alignReal = value;
                HasAlignment = true;
            }
        }

        internal bool HasProtection;
        private SLProtection protectionReal;
        /// <summary>
        /// The protection settings for incremental formatting.
        /// </summary>
        public SLProtection Protection
        {
            get { return protectionReal; }
            set
            {
                protectionReal = value;
                HasProtection = true;
            }
        }

        internal bool HasNumberingFormat;
        internal SLNumberingFormat nfFormatCode;
        /// <summary>
        /// The numbering format for incremental formatting.
        /// </summary>
        public string FormatCode
        {
            get { return nfFormatCode.FormatCode; }
            set
            {
                nfFormatCode.FormatCode = value.Trim();
                if (nfFormatCode.FormatCode.Length > 0)
                {
                    HasNumberingFormat = true;
                }
                else
                {
                    HasNumberingFormat = false;
                }
            }
        }

        internal bool HasFont;
        private SLFont fontReal;
        /// <summary>
        /// The font for incremental formatting.
        /// </summary>
        public SLFont Font
        {
            get { return fontReal; }
            set
            {
                fontReal = value;
                HasFont = true;
            }
        }

        internal bool HasFill;
        private SLFill fillReal;
        /// <summary>
        /// The fill for incremental formatting.
        /// </summary>
        public SLFill Fill
        {
            get { return fillReal; }
            set
            {
                fillReal = value;
                HasFill = true;
            }
        }

        internal bool HasBorder;
        private SLBorder borderReal;
        /// <summary>
        /// The border for incremental formatting.
        /// </summary>
        public SLBorder Border
        {
            get { return borderReal; }
            set
            {
                borderReal = value;
                HasBorder = true;
            }
        }

        /// <summary>
        /// Initializes an instance of SLDifferentialFormat.
        /// </summary>
        public SLDifferentialFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();

            this.alignReal = new SLAlignment();
            HasAlignment = false;
            this.protectionReal = new SLProtection();
            HasProtection = false;
            this.nfFormatCode = new SLNumberingFormat();
            HasNumberingFormat = false;
            this.fontReal = new SLFont(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, listempty, listempty);
            HasFont = false;
            this.fillReal = new SLFill(listempty, listempty);
            HasFill = false;
            this.borderReal = new SLBorder(listempty, listempty);
            HasBorder = false;
        }

        internal void Sync()
        {
            HasAlignment = Alignment.HasHorizontal || Alignment.HasVertical || Alignment.TextRotation != null || Alignment.WrapText != null || Alignment.Indent != null || Alignment.RelativeIndent != null || Alignment.JustifyLastLine != null || Alignment.ShrinkToFit != null || Alignment.HasReadingOrder;
            HasProtection = Protection.Locked != null || Protection.Hidden != null;
            //HasNumberingFormat
            HasFont = Font.FontName != null || Font.CharacterSet != null || Font.FontFamily != null || Font.Bold != null || Font.Italic != null || Font.Strike != null || Font.Outline != null || Font.Shadow != null || Font.Condense != null || Font.Extend != null || Font.HasFontColor || Font.FontSize != null || Font.HasUnderline || Font.HasVerticalAlignment || Font.HasFontScheme;
            HasFill = Fill.HasBeenAssignedValues;
            Border.Sync();
            HasBorder = Border.HasLeftBorder || Border.HasRightBorder || Border.HasTopBorder || Border.HasBottomBorder || Border.HasDiagonalBorder || Border.HasVerticalBorder || Border.HasHorizontalBorder || Border.DiagonalUp != null || Border.DiagonalDown != null || Border.Outline != null;
        }

        internal void FromDifferentialFormat(DifferentialFormat df)
        {
            this.SetAllNull();

            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();

            if (df.Font != null)
            {
                HasFont = true;
                this.fontReal = new SLFont(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, listempty, listempty);
                this.fontReal.FromFont(df.Font);
            }

            if (df.NumberingFormat != null)
            {
                HasNumberingFormat = true;
                this.nfFormatCode = new SLNumberingFormat();
                this.nfFormatCode.FromNumberingFormat(df.NumberingFormat);
            }

            if (df.Fill != null)
            {
                HasFill = true;
                this.fillReal = new SLFill(listempty, listempty);
                this.fillReal.FromFill(df.Fill);
            }

            if (df.Alignment != null)
            {
                HasAlignment = true;
                this.alignReal = new SLAlignment();
                this.alignReal.FromAlignment(df.Alignment);
            }

            if (df.Border != null)
            {
                HasBorder = true;
                this.borderReal = new SLBorder(listempty, listempty);
                this.borderReal.FromBorder(df.Border);
            }

            if (df.Protection != null)
            {
                HasProtection = true;
                this.protectionReal = new SLProtection();
                this.protectionReal.FromProtection(df.Protection);
            }

            Sync();
        }

        internal void FromDifferentialType(X14.DifferentialType dt)
        {
            this.SetAllNull();

            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();

            if (dt.Font != null)
            {
                HasFont = true;
                this.fontReal = new SLFont(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, listempty, listempty);
                this.fontReal.FromFont(dt.Font);
            }

            if (dt.NumberingFormat != null)
            {
                HasNumberingFormat = true;
                this.nfFormatCode = new SLNumberingFormat();
                this.nfFormatCode.FromNumberingFormat(dt.NumberingFormat);
            }

            if (dt.Fill != null)
            {
                HasFill = true;
                this.fillReal = new SLFill(listempty, listempty);
                this.fillReal.FromFill(dt.Fill);
            }

            if (dt.Alignment != null)
            {
                HasAlignment = true;
                this.alignReal = new SLAlignment();
                this.alignReal.FromAlignment(dt.Alignment);
            }

            if (dt.Border != null)
            {
                HasBorder = true;
                this.borderReal = new SLBorder(listempty, listempty);
                this.borderReal.FromBorder(dt.Border);
            }

            if (dt.Protection != null)
            {
                HasProtection = true;
                this.protectionReal = new SLProtection();
                this.protectionReal.FromProtection(dt.Protection);
            }

            Sync();
        }

        internal DifferentialFormat ToDifferentialFormat()
        {
            Sync();

            DifferentialFormat df = new DifferentialFormat();
            if (HasFont) df.Font = this.Font.ToFont();
            if (HasNumberingFormat) df.NumberingFormat = this.nfFormatCode.ToNumberingFormat();
            if (HasFill) df.Fill = this.Fill.ToFill();
            if (HasAlignment) df.Alignment = this.Alignment.ToAlignment();
            if (HasBorder) df.Border = this.Border.ToBorder();
            if (HasProtection) df.Protection = this.Protection.ToProtection();

            return df;
        }

        internal X14.DifferentialType ToDifferentialType()
        {
            Sync();

            X14.DifferentialType dt = new X14.DifferentialType();
            if (HasFont) dt.Font = this.Font.ToFont();
            if (HasNumberingFormat) dt.NumberingFormat = this.nfFormatCode.ToNumberingFormat();
            if (HasFill) dt.Fill = this.Fill.ToFill();
            if (HasAlignment) dt.Alignment = this.Alignment.ToAlignment();
            if (HasBorder) dt.Border = this.Border.ToBorder();
            if (HasProtection) dt.Protection = this.Protection.ToProtection();

            return dt;
        }

        internal void FromHash(string Hash)
        {
            DifferentialFormat df = new DifferentialFormat();
            df.InnerXml = Hash;
            this.FromDifferentialFormat(df);
        }

        internal string ToHash()
        {
            DifferentialFormat df = this.ToDifferentialFormat();
            return SLTool.RemoveNamespaceDeclaration(df.InnerXml);
        }

        internal SLDifferentialFormat Clone()
        {
            SLDifferentialFormat df = new SLDifferentialFormat();
            df.HasAlignment = this.HasAlignment;
            df.alignReal = this.alignReal.Clone();
            df.HasProtection = this.HasProtection;
            df.protectionReal = this.protectionReal.Clone();
            df.HasNumberingFormat = this.HasNumberingFormat;
            df.nfFormatCode = this.nfFormatCode.Clone();
            df.HasFont = this.HasFont;
            df.fontReal = this.fontReal.Clone();
            df.HasFill = this.HasFill;
            df.fillReal = this.fillReal.Clone();
            df.HasBorder = this.HasBorder;
            df.borderReal = this.borderReal.Clone();

            return df;
        }
    }
}
