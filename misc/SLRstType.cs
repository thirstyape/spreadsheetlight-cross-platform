using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for handling rich string types. This includes the CommentText class, InlineString class and SharedStringItem class. This simulates the (abstract) DocumentFormat.OpenXml.Spreadsheet.RstType class.
    /// </summary>
    /// <remarks>This also take on double duty as rich text for other purposes such as charts. We do this so other developers don't have to learn another class.</remarks>
    public class SLRstType
    {
        internal string MajorFont { get; set; }
        internal string MinorFont { get; set; }
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        internal InlineString istrReal;

        /// <summary>
        /// Initializes an instance of SLRstType. It is recommended to use CreateRstType() of the SLDocument class.
        /// </summary>
        public SLRstType()
        {
            this.Initialize(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
        }

        internal SLRstType(string MajorFont, string MinorFont, List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            this.Initialize(MajorFont, MinorFont, ThemeColors, IndexedColors);
        }

        private void Initialize(string MajorFont, string MinorFont, List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            this.MajorFont = MajorFont;
            this.MinorFont = MinorFont;

            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }

            this.listIndexedColors = new List<System.Drawing.Color>();
            for (i = 0; i < IndexedColors.Count; ++i)
            {
                this.listIndexedColors.Add(IndexedColors[i]);
            }

            istrReal = new InlineString();
        }

        /// <summary>
        /// Set the text. If text formatting is needed, use one of the AppendText() functions.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetText(string Text)
        {
            if (Text == null || Text.Length == 0)
            {
                this.istrReal.Text = null;
            }
            else
            {
                this.istrReal.Text = new Text();
                this.istrReal.Text.Text = Text;
                if (SLTool.ToPreserveSpace(Text))
                {
                    this.istrReal.Text.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }

        /// <summary>
        /// Get the text. This is the plain text string, and not the rich text runs.
        /// </summary>
        /// <returns>The plain text.</returns>
        public string GetText()
        {
            string result = string.Empty;
            if (this.istrReal.Text != null) result = this.istrReal.Text.Text;

            return result;
        }

        /// <summary>
        /// Get a list of rich text runs.
        /// </summary>
        /// <returns>A list of rich text runs.</returns>
        public List<SLRun> GetRuns()
        {
            List<SLRun> result = new List<SLRun>();
            SLRun r;

            using (OpenXmlReader oxr = OpenXmlReader.Create(this.istrReal))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Run))
                    {
                        r = new SLRun();
                        r.FromRun((Run)oxr.LoadCurrentElement());
                        result.Add(r.Clone());
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Replace the internal list of rich text runs.
        /// </summary>
        /// <param name="Runs">The new list of rich text runs for replacing.</param>
        public void ReplaceRuns(List<SLRun> Runs)
        {
            string sText = string.Empty;
            if (this.istrReal.Text != null) sText = this.istrReal.Text.Text;

            this.istrReal.RemoveAllChildren<Text>();
            this.istrReal.RemoveAllChildren<Run>();

            int i;
            // start from the end because we're prepending to the front
            for (i = Runs.Count - 1; i >= 0; --i)
            {
                this.istrReal.PrependChild<Run>(Runs[i].ToRun());
            }

            this.SetText(sText);
        }

        /// <summary>
        /// Append given text in the current theme's minor font and default font size.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void AppendText(string Text)
        {
            SLFont font = new SLFont(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
            font.SetFont(FontSchemeValues.Minor, SLConstants.DefaultFontSize);

            this.AppendText(Text, font);
        }

        /// <summary>
        /// Append given text with a given font style.
        /// </summary>
        /// <param name="Text">The text.</param>
        /// <param name="TextFont">The font style.</param>
        public void AppendText(string Text, SLFont TextFont)
        {
            Run run = new Run();
            RunProperties runprops = new RunProperties();

            if (TextFont.FontName != null)
            {
                runprops.Append(new RunFont() { Val = TextFont.FontName });
            }

            if (TextFont.CharacterSet != null)
            {
                runprops.Append(new RunPropertyCharSet() { Val = TextFont.CharacterSet.Value });
            }

            if (TextFont.FontFamily != null)
            {
                runprops.Append(new FontFamily() { Val = TextFont.FontFamily.Value });
            }

            if (TextFont.Bold != null)
            {
                runprops.Append(new Bold() { Val = TextFont.Bold.Value });
            }

            if (TextFont.Italic != null)
            {
                runprops.Append(new Italic() { Val = TextFont.Italic.Value });
            }

            if (TextFont.Strike != null)
            {
                runprops.Append(new Strike() { Val = TextFont.Strike.Value });
            }

            if (TextFont.Outline != null)
            {
                runprops.Append(new Outline() { Val = TextFont.Outline.Value });
            }

            if (TextFont.Shadow != null)
            {
                runprops.Append(new Shadow() { Val = TextFont.Shadow.Value });
            }

            if (TextFont.Condense != null)
            {
                runprops.Append(new Condense() { Val = TextFont.Condense.Value });
            }

            if (TextFont.Extend != null)
            {
                runprops.Append(new Extend() { Val = TextFont.Extend.Value });
            }

            if (TextFont.HasFontColor)
            {
                runprops.Append(TextFont.clrFontColor.ToSpreadsheetColor());
            }

            if (TextFont.FontSize != null)
            {
                runprops.Append(new FontSize() { Val = TextFont.FontSize.Value });
            }

            if (TextFont.HasUnderline)
            {
                runprops.Append(new Underline() { Val = TextFont.Underline });
            }

            if (TextFont.HasVerticalAlignment)
            {
                runprops.Append(new VerticalTextAlignment() { Val = TextFont.VerticalAlignment });
            }

            if (TextFont.HasFontScheme)
            {
                runprops.Append(new FontScheme() { Val = TextFont.FontScheme });
            }

            if (runprops.ChildElements.Count > 0)
            {
                run.Append(runprops);
            }

            run.Text = new Text();
            run.Text.Text = Text;
            if (SLTool.ToPreserveSpace(Text))
            {
                run.Text.Space = SpaceProcessingModeValues.Preserve;
            }

            bool bFound = false;
            OpenXmlElement oxe = istrReal.FirstChild;
            foreach (var child in this.istrReal.ChildElements)
            {
                if (child is Text || child is Run)
                {
                    oxe = child;
                    bFound = true;
                }
            }

            if (bFound)
            {
                istrReal.InsertAfter(run, oxe);
            }
            else
            {
                istrReal.PrependChild(run);
            }
        }

        /// <summary>
        /// Form an SLRstType from DocumentFormat.OpenXml.Spreadsheet.CommentText class.
        /// </summary>
        /// <param name="Comment">A source DocumentFormat.OpenXml.Spreadsheet.CommentText class.</param>
        public void FromCommentText(CommentText Comment)
        {
            this.istrReal.InnerXml = Comment.InnerXml;
        }

        /// <summary>
        /// Form a DocumentFormat.OpenXml.Spreadsheet.CommentText class from this SLRstType class.
        /// </summary>
        /// <returns>A DocumentFormat.OpenXml.Spreadsheet.CommentText class.</returns>
        public CommentText ToCommentText()
        {
            CommentText ct = new CommentText();
            ct.InnerXml = SLTool.RemoveNamespaceDeclaration(this.istrReal.InnerXml);
            return ct;
        }

        /// <summary>
        /// Form an SLRstType from DocumentFormat.OpenXml.Spreadsheet.InlineString class.
        /// </summary>
        /// <param name="RichText">A source DocumentFormat.OpenXml.Spreadsheet.InlineString class.</param>
        public void FromInlineString(InlineString RichText)
        {
            this.istrReal.InnerXml = RichText.InnerXml;
        }

        /// <summary>
        /// Form a DocumentFormat.OpenXml.Spreadsheet.InlineString class from this SLRstType class.
        /// </summary>
        /// <returns>A DocumentFormat.OpenXml.Spreadsheet.InlineString class.</returns>
        public InlineString ToInlineString()
        {
            InlineString istr = new InlineString();
            istr.InnerXml = SLTool.RemoveNamespaceDeclaration(this.istrReal.InnerXml);
            return istr;
        }

        /// <summary>
        /// Form an SLRstType from DocumentFormat.OpenXml.Spreadsheet.SharedStringItem class.
        /// </summary>
        /// <param name="SharedString">A source DocumentFormat.OpenXml.Spreadsheet.SharedStringItem class.</param>
        public void FromSharedStringItem(SharedStringItem SharedString)
        {
            this.istrReal.InnerXml = SharedString.InnerXml;
        }

        /// <summary>
        /// Form a DocumentFormat.OpenXml.Spreadsheet.SharedStringItem class from this SLRstType class.
        /// </summary>
        /// <returns>A DocumentFormat.OpenXml.Spreadsheet.SharedStringItem class.</returns>
        public SharedStringItem ToSharedStringItem()
        {
            SharedStringItem ssi = new SharedStringItem();
            ssi.InnerXml = SLTool.RemoveNamespaceDeclaration(this.istrReal.InnerXml);
            return ssi;
        }

        /// <summary>
        /// Form a string with all the text formatting stripped out.
        /// </summary>
        /// <returns>An unformatted text string.</returns>
        public string ToPlainString()
        {
            StringBuilder sb = new StringBuilder();
            if (this.istrReal.Text != null)
            {
                sb.Append(this.istrReal.Text.Text);
            }

            Run run;
            //PhoneticRun phrun;
            foreach (var child in this.istrReal.ChildElements)
            {
                if (child is Run)
                {
                    run = (Run)child;
                    // the Text child class is compulsory, but just in case...
                    if (run.Text != null)
                    {
                        sb.Append(run.Text.Text);
                    }
                }

                // 19 Oct 2016: No need to retrieve phonetic text. I don't know how it's
                // used. I don't type Japanese and other phonetic-related languages.
                // Apparently this is auto-entered by Excel when you type Japanese into
                // the cells (via Excel, not via libraries).
                // Thanks to Vincent D (yes, same first name :) for pointing this out.
                //else if (child is PhoneticRun)
                //{
                //    phrun = (PhoneticRun)child;
                //    // the Text child class is compulsory, but just in case...
                //    if (phrun.Text != null)
                //    {
                //        sb.Append(phrun.Text.Text);
                //    }
                //}
            }

            return sb.ToString();
        }

        internal void FromHash(string Hash)
        {
            InlineString istr = new InlineString();
            istr.InnerXml = Hash;
            this.FromInlineString(istr);
        }

        internal string ToHash()
        {
            InlineString istr = this.ToInlineString();
            return SLTool.RemoveNamespaceDeclaration(istr.InnerXml);
        }

        internal A.Paragraph ToParagraph()
        {
            A.Paragraph para = new A.Paragraph();
            para.ParagraphProperties = new A.ParagraphProperties();
            para.ParagraphProperties.Append(new A.DefaultRunProperties());

            A.Run run;

            if (this.istrReal.Text != null)
            {
                run = new A.Run();
                run.RunProperties = new A.RunProperties();
                run.Text = new A.Text(this.istrReal.Text.Text);
                para.Append(run);
            }

            Run xrun;

            RunFont xrunRunFont;
            Bold xrunBold;
            Italic xrunItalic;
            Strike xrunStrike;
            Color xrunColor;
            FontSize xrunFontSize;
            Underline xrunUnderline;
            VerticalTextAlignment xrunVertical;
            FontScheme xrunScheme;

            string sFont;
            bool? bBold;
            bool? bItalic;
            bool? bStrike;
            double? fFontSize;
            UnderlineValues? vUnderline;
            VerticalAlignmentRunValues? vVertical;
            bool bHasColor;
            SLA.SLColorTransform clrRun = new SLA.SLColorTransform(new List<System.Drawing.Color>());
            FontSchemeValues? vScheme;

            using (OpenXmlReader oxr = OpenXmlReader.Create(this.istrReal))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Run))
                    {
                        run = new A.Run();
                        run.RunProperties = new A.RunProperties();

                        sFont = string.Empty;
                        bBold = null;
                        bItalic = null;
                        bStrike = null;
                        fFontSize = null;
                        vUnderline = null;
                        vVertical = null;
                        bHasColor = false;
                        vScheme = null;

                        xrun = (Run)oxr.LoadCurrentElement();
                        if (xrun.RunProperties != null)
                        {
                            using (OpenXmlReader oxrProps = OpenXmlReader.Create(xrun.RunProperties))
                            {
                                while (oxrProps.Read())
                                {
                                    if (oxrProps.ElementType == typeof(RunFont))
                                    {
                                        xrunRunFont = (RunFont)oxrProps.LoadCurrentElement();
                                        if (xrunRunFont.Val != null) sFont = xrunRunFont.Val.Value;
                                    }
                                    else if (oxrProps.ElementType == typeof(Bold))
                                    {
                                        xrunBold = (Bold)oxrProps.LoadCurrentElement();
                                        if (xrunBold.Val != null) bBold = xrunBold.Val.Value;
                                    }
                                    else if (oxrProps.ElementType == typeof(Italic))
                                    {
                                        xrunItalic = (Italic)oxrProps.LoadCurrentElement();
                                        if (xrunItalic.Val != null) bItalic = xrunItalic.Val.Value;
                                    }
                                    else if (oxrProps.ElementType == typeof(Strike))
                                    {
                                        xrunStrike = (Strike)oxrProps.LoadCurrentElement();
                                        if (xrunStrike.Val != null) bStrike = xrunStrike.Val.Value;
                                    }
                                    else if (oxrProps.ElementType == typeof(Color))
                                    {
                                        xrunColor = (Color)oxrProps.LoadCurrentElement();
                                        if (xrunColor.Rgb != null)
                                        {
                                            bHasColor = true;
                                            clrRun = new SLA.SLColorTransform(new List<System.Drawing.Color>());
                                            clrRun.SetColor(SLTool.ToColor(xrunColor.Rgb.Value), 0);
                                        }
                                        else if (xrunColor.Theme != null)
                                        {
                                            bHasColor = true;
                                            clrRun = new SLA.SLColorTransform(new List<System.Drawing.Color>());
                                            if (xrunColor.Tint != null)
                                            {
                                                clrRun.SetColor((SLThemeColorIndexValues)xrunColor.Theme.Value, xrunColor.Tint.Value, 0);
                                            }
                                            else
                                            {
                                                clrRun.SetColor((SLThemeColorIndexValues)xrunColor.Theme.Value, 0, 0);
                                            }
                                        }
                                    }
                                    else if (oxrProps.ElementType == typeof(FontSize))
                                    {
                                        xrunFontSize = (FontSize)oxrProps.LoadCurrentElement();
                                        if (xrunFontSize.Val != null) fFontSize = xrunFontSize.Val.Value;
                                    }
                                    else if (oxrProps.ElementType == typeof(Underline))
                                    {
                                        xrunUnderline = (Underline)oxrProps.LoadCurrentElement();
                                        if (xrunUnderline.Val != null) vUnderline = xrunUnderline.Val.Value;
                                    }
                                    else if (oxrProps.ElementType == typeof(VerticalTextAlignment))
                                    {
                                        xrunVertical = (VerticalTextAlignment)oxrProps.LoadCurrentElement();
                                        if (xrunVertical.Val != null) vVertical = xrunVertical.Val.Value;
                                    }
                                    else if (oxrProps.ElementType == typeof(FontScheme))
                                    {
                                        xrunScheme = (FontScheme)oxrProps.LoadCurrentElement();
                                        if (xrunScheme.Val != null) vScheme = xrunScheme.Val.Value;
                                    }
                                }
                            }
                        }

                        if (vScheme != null)
                        {
                            if (vScheme.Value == FontSchemeValues.Major) sFont = "+mj-lt";
                            else if (vScheme.Value == FontSchemeValues.Minor) sFont = "+mn-lt";
                        }

                        if (bHasColor)
                        {
                            if (clrRun.IsRgbColorModelHex)
                            {
                                run.RunProperties.Append(new A.SolidFill()
                                {
                                    RgbColorModelHex = clrRun.ToRgbColorModelHex()
                                });
                            }
                            else
                            {
                                run.RunProperties.Append(new A.SolidFill()
                                {
                                    SchemeColor = clrRun.ToSchemeColor()
                                });
                            }
                        }

                        if (sFont.Length > 0) run.RunProperties.Append(new A.LatinFont() { Typeface = sFont });

                        if (fFontSize != null) run.RunProperties.FontSize = (int)(fFontSize.Value * 100);

                        if (bBold != null) run.RunProperties.Bold = bBold.Value;

                        if (bItalic != null) run.RunProperties.Italic = bItalic.Value;

                        if (vUnderline != null)
                        {
                            if (vUnderline.Value == UnderlineValues.Single || vUnderline.Value == UnderlineValues.SingleAccounting)
                            {
                                run.RunProperties.Underline = A.TextUnderlineValues.Single;
                            }
                            else if (vUnderline.Value == UnderlineValues.Double || vUnderline.Value == UnderlineValues.DoubleAccounting)
                            {
                                run.RunProperties.Underline = A.TextUnderlineValues.Double;
                            }
                        }

                        if (bStrike != null)
                        {
                            run.RunProperties.Strike = bStrike.Value ? A.TextStrikeValues.SingleStrike : A.TextStrikeValues.NoStrike;
                        }

                        if (vVertical != null)
                        {
                            if (vVertical.Value == VerticalAlignmentRunValues.Superscript)
                            {
                                run.RunProperties.Baseline = 30000;
                            }
                            else if (vVertical.Value == VerticalAlignmentRunValues.Subscript)
                            {
                                run.RunProperties.Baseline = -25000;
                            }
                            else
                            {
                                run.RunProperties.Baseline = 0;
                            }
                        }
                        else
                        {
                            run.RunProperties.Baseline = 0;
                        }

                        run.Text = new A.Text(xrun.Text.Text);
                        para.Append(run);
                    }
                }
            }

            return para;
        }

        /// <summary>
        /// Clone a new instance of SLRstType.
        /// </summary>
        /// <returns>A cloned instance of this SLRstType.</returns>
        public SLRstType Clone()
        {
            SLRstType rst = new SLRstType(this.MajorFont, this.MinorFont, this.listThemeColors, this.listIndexedColors);
            rst.istrReal = (InlineString)this.istrReal.CloneNode(true);

            return rst;
        }
    }
}
