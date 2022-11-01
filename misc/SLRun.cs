using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for rich text runs. This simulates the DocumentFormat.OpenXml.Spreadsheet.Run class.
    /// </summary>
    public class SLRun
    {
        /// <summary>
        /// The font styles.
        /// </summary>
        public SLFont Font { get; set; }

        /// <summary>
        /// The text.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Initializes an instance of SLRun.
        /// </summary>
        public SLRun()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Font = new SLFont(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.Text = string.Empty;
        }

        internal void FromRun(Run r)
        {
            this.SetAllNull();

            using (OpenXmlReader oxr = OpenXmlReader.Create(r))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Text))
                    {
                        this.Text = ((Text)oxr.LoadCurrentElement()).Text;
                    }
                    else if (oxr.ElementType == typeof(RunFont))
                    {
                        RunFont rft = (RunFont)oxr.LoadCurrentElement();
                        if (rft.Val != null) this.Font.FontName = rft.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(RunPropertyCharSet))
                    {
                        RunPropertyCharSet rpcs = (RunPropertyCharSet)oxr.LoadCurrentElement();
                        if (rpcs.Val != null) this.Font.CharacterSet = rpcs.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(FontFamily))
                    {
                        FontFamily ff = (FontFamily)oxr.LoadCurrentElement();
                        if (ff.Val != null) this.Font.FontFamily = ff.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(Bold))
                    {
                        Bold b = (Bold)oxr.LoadCurrentElement();
                        if (b.Val != null) this.Font.Bold = b.Val.Value;
                        else this.Font.Bold = true;
                    }
                    else if (oxr.ElementType == typeof(Italic))
                    {
                        Italic itlc = (Italic)oxr.LoadCurrentElement();
                        if (itlc.Val != null) this.Font.Italic = itlc.Val.Value;
                        else this.Font.Italic = true;
                    }
                    else if (oxr.ElementType == typeof(Strike))
                    {
                        Strike strk = (Strike)oxr.LoadCurrentElement();
                        if (strk.Val != null) this.Font.Strike = strk.Val.Value;
                        else this.Font.Strike = true;
                    }
                    else if (oxr.ElementType == typeof(Outline))
                    {
                        Outline outln = (Outline)oxr.LoadCurrentElement();
                        if (outln.Val != null) this.Font.Outline = outln.Val.Value;
                        else this.Font.Outline = true;
                    }
                    else if (oxr.ElementType == typeof(Shadow))
                    {
                        Shadow shdw = (Shadow)oxr.LoadCurrentElement();
                        if (shdw.Val != null) this.Font.Shadow = shdw.Val.Value;
                        else this.Font.Shadow = true;
                    }
                    else if (oxr.ElementType == typeof(Condense))
                    {
                        Condense cdns = (Condense)oxr.LoadCurrentElement();
                        if (cdns.Val != null) this.Font.Condense = cdns.Val.Value;
                        else this.Font.Condense = true;
                    }
                    else if (oxr.ElementType == typeof(Extend))
                    {
                        Extend ext = (Extend)oxr.LoadCurrentElement();
                        if (ext.Val != null) this.Font.Extend = ext.Val.Value;
                        else this.Font.Extend = true;
                    }
                    else if (oxr.ElementType == typeof(Color))
                    {
                        this.Font.clrFontColor.FromSpreadsheetColor((Color)oxr.LoadCurrentElement());
                        this.Font.HasFontColor = !this.Font.clrFontColor.IsEmpty();
                    }
                    else if (oxr.ElementType == typeof(FontSize))
                    {
                        FontSize ftsz = (FontSize)oxr.LoadCurrentElement();
                        if (ftsz.Val != null) this.Font.FontSize = ftsz.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(Underline))
                    {
                        Underline und = (Underline)oxr.LoadCurrentElement();
                        if (und.Val != null) this.Font.Underline = und.Val.Value;
                        else this.Font.Underline = UnderlineValues.Single;
                    }
                    else if (oxr.ElementType == typeof(VerticalTextAlignment))
                    {
                        VerticalTextAlignment vta = (VerticalTextAlignment)oxr.LoadCurrentElement();
                        if (vta.Val != null) this.Font.VerticalAlignment = vta.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(FontScheme))
                    {
                        FontScheme ftsch = (FontScheme)oxr.LoadCurrentElement();
                        if (ftsch.Val != null) this.Font.FontScheme = ftsch.Val.Value;
                    }
                }
            }
        }

        internal Run ToRun()
        {
            Run r = new Run();
            r.RunProperties = new RunProperties();

            if (this.Font.FontName != null)
            {
                r.RunProperties.Append(new RunFont() { Val = this.Font.FontName });
            }

            if (this.Font.CharacterSet != null)
            {
                r.RunProperties.Append(new RunPropertyCharSet() { Val = this.Font.CharacterSet.Value });
            }

            if (this.Font.FontFamily != null)
            {
                r.RunProperties.Append(new FontFamily() { Val = this.Font.FontFamily.Value });
            }

            if (this.Font.Bold != null && this.Font.Bold.Value)
            {
                r.RunProperties.Append(new Bold());
            }

            if (this.Font.Italic != null && this.Font.Italic.Value)
            {
                r.RunProperties.Append(new Italic());
            }

            if (this.Font.Strike != null && this.Font.Strike.Value)
            {
                r.RunProperties.Append(new Strike());
            }

            if (this.Font.Outline != null && this.Font.Outline.Value)
            {
                r.RunProperties.Append(new Outline());
            }

            if (this.Font.Shadow != null && this.Font.Shadow.Value)
            {
                r.RunProperties.Append(new Shadow());
            }

            if (this.Font.Condense != null && this.Font.Condense.Value)
            {
                r.RunProperties.Append(new Condense());
            }

            if (this.Font.Extend != null && this.Font.Extend.Value)
            {
                r.RunProperties.Append(new Extend());
            }

            if (this.Font.HasFontColor)
            {
                r.RunProperties.Append(this.Font.clrFontColor.ToSpreadsheetColor());
            }

            if (this.Font.FontSize != null)
            {
                r.RunProperties.Append(new FontSize() { Val = this.Font.FontSize.Value });
            }

            if (this.Font.HasUnderline)
            {
                r.RunProperties.Append(new Underline() { Val = this.Font.Underline });
            }

            if (this.Font.HasVerticalAlignment)
            {
                r.RunProperties.Append(new VerticalTextAlignment() { Val = this.Font.VerticalAlignment });
            }

            if (this.Font.HasFontScheme)
            {
                r.RunProperties.Append(new FontScheme() { Val = this.Font.FontScheme });
            }

            r.Text = new Text(this.Text);
            if (SLTool.ToPreserveSpace(this.Text)) r.Text.Space = SpaceProcessingModeValues.Preserve;

            return r;
        }

        /// <summary>
        /// Clone a new instance of SLRun.
        /// </summary>
        /// <returns>An SLRun object.</returns>
        public SLRun Clone()
        {
            SLRun r = new SLRun();
            r.Font = this.Font.Clone();
            r.Text = this.Text;

            return r;
        }
    }
}
