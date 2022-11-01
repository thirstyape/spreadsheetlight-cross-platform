using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal enum SLHeaderFooterSection
    {
        None = 0,
        OddHeader,
        OddFooter,
        EvenHeader,
        EvenFooter,
        FirstHeader,
        FirstFooter
    }

    /// <summary>
    /// Encapsulates page and print settings for a sheet (worksheets, chartsheets and dialogsheets).
    /// This simulates DocumentFormat.OpenXml.Spreadsheet.SheetProperties, DocumentFormat.OpenXml.Spreadsheet.PrintOptions, DocumentFormat.OpenXml.Spreadsheet.PageMargins,
    /// DocumentFormat.OpenXml.Spreadsheet.PageSetup, DocumentFormat.OpenXml.Spreadsheet.HeaderFooter and DocumentFormat.OpenXml.Spreadsheet.SheetView classes.
    /// For chartsheets, the DocumentFormat.OpenXml.Spreadsheet.ChartSheetProperties (instead of SheetProperties) and
    /// DocumentFormat.OpenXml.Spreadsheet.ChartSheetPageSetup (instead of PageSetup) classes are involved.
    /// </summary>
    public class SLPageSettings
    {
        //SheetProperties: TabColor
        //SheetProperties: PageSetupProperties
        //SheetViews: Zoom
        //PrintOptions (parents: customSheetView, dialogsheet, worksheet)
        //PageMargins (parents: chartsheet, customSheetView, dialogsheet, worksheet)
        //PageSetup (parents: customSheetView, dialogsheet, worksheet)
        //HeaderFooter (parents: chartsheet, customSheetView, dialogsheet, worksheet)

        internal bool HasSheetProperties
        {
            get
            {
                return this.SheetProperties.HasSheetProperties;
            }
        }

        internal bool HasChartSheetProperties
        {
            get
            {
                return this.SheetProperties.HasChartSheetProperties;
            }
        }

        internal SLSheetProperties SheetProperties { get; set; }

        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        /// <summary>
        /// Specifies if there's a tab color. This is read-only.
        /// </summary>
        public bool HasTabColor
        {
            get
            {
                return this.SheetProperties.HasTabColor;
            }
        }

        /// <summary>
        /// The tab color.
        /// </summary>
        public System.Drawing.Color TabColor
        {
            get { return this.SheetProperties.clrTabColor.Color; }
            set
            {
                this.SheetProperties.TabColor = value;
            }
        }

        internal bool HasSheetView
        {
            get
            {
                return this.bShowFormulas != null || this.bShowGridLines != null
                    || this.bShowRowColumnHeaders != null || this.bShowRuler != null
                    || this.vView != null || this.iZoomScale != null
                    || this.iZoomScaleNormal != null || this.iZoomScalePageLayoutView != null;
            }
        }

        // Apparently when show formulas is true, column widths are doubled. Erhmahgerd...

        internal bool? bShowFormulas;
        /// <summary>
        /// Show or hide the cell formulas. NOTE: This has nothing to do with the formula bar, but whether the sheet shows cell formulas instead of calculated results.
        /// </summary>
        public bool ShowFormulas
        {
            get { return bShowFormulas ?? false; }
            set { bShowFormulas = value; }
        }

        internal bool? bShowGridLines;
        /// <summary>
        /// Show or hide the grid lines between rows and columns.
        /// </summary>
        public bool ShowGridLines
        {
            get { return bShowGridLines ?? true; }
            set { bShowGridLines = value; }
        }

        internal bool? bShowRowColumnHeaders;
        /// <summary>
        /// Show or hide the row and column headers.
        /// </summary>
        public bool ShowRowColumnHeaders
        {
            get { return bShowRowColumnHeaders ?? true; }
            set { bShowRowColumnHeaders = value; }
        }

        internal bool? bShowRuler;
        /// <summary>
        /// Show or hide the ruler on the worksheet. The ruler is only seen when the worksheet view is in "page layout" mode.
        /// </summary>
        public bool ShowRuler
        {
            get { return bShowRuler ?? true; }
            set { bShowRuler = value; }
        }

        internal SheetViewValues? vView;
        /// <summary>
        /// Worksheet view type.
        /// </summary>
        public SheetViewValues View
        {
            get { return vView ?? SheetViewValues.Normal; }
            set { vView = value; }
        }

        internal uint? iZoomScale;
        /// <summary>
        /// Zoom magnification for current view, ranging from 10% to 400%. If you want to set a zoom value for the page break view, make sure to set the View property to PageBreakPreview.
        /// </summary>
        public uint ZoomScale
        {
            get { return iZoomScale ?? 100; }
            set
            {
                iZoomScale = value;
                if (iZoomScale < 10) iZoomScale = 10;
                if (iZoomScale > 400) iZoomScale = 400;
            }
        }

        internal uint? iZoomScaleNormal;
        /// <summary>
        /// Zoom magnification for the normal view, ranging from 10% to 400%. A return value of 0% means the automatic setting is used.
        /// If the view is set to normal, this value is ignored if ZoomScale is also set.
        /// </summary>
        public uint ZoomScaleNormal
        {
            get { return iZoomScaleNormal ?? 0; }
            set
            {
                iZoomScaleNormal = value;
                if (iZoomScaleNormal < 10) iZoomScaleNormal = 10;
                if (iZoomScaleNormal > 400) iZoomScaleNormal = 400;
            }
        }

        internal uint? iZoomScalePageLayoutView;
        /// <summary>
        /// Zoom magnification for the page layout view, ranging from 10% to 400%. A return value of 0% means the automatic setting is used.
        /// If the view is set to page layout, this value is ignored if ZoomScale is also set.
        /// </summary>
        public uint ZoomScalePageLayoutView
        {
            get { return iZoomScalePageLayoutView ?? 0; }
            set
            {
                iZoomScalePageLayoutView = value;
                if (iZoomScalePageLayoutView < 10) iZoomScalePageLayoutView = 10;
                if (iZoomScalePageLayoutView > 400) iZoomScalePageLayoutView = 400;
            }
        }

        internal bool HasPrintOptions
        {
            get
            {
                return PrintHorizontalCentered || PrintVerticalCentered || PrintHeadings || PrintGridLines;
            }
        }

        /// <summary>
        /// Center horizontally on page when printing. This doesn't apply to chart sheets.
        /// </summary>
        public bool PrintHorizontalCentered { get; set; }

        /// <summary>
        /// Center vertically on page when printing. This doesn't apply to chart sheets.
        /// </summary>
        public bool PrintVerticalCentered { get; set; }

        /// <summary>
        /// Print row and column headings. This doesn't apply to chart sheets.
        /// </summary>
        public bool PrintHeadings { get; set; }

        /// <summary>
        /// Print grid lines. This doesn't apply to chart sheets.
        /// </summary>
        public bool PrintGridLines { get; set; }

        internal bool PrintGridLinesSet { get; set; }

        internal bool HasPageMargins;

        internal double fLeftMargin;
        /// <summary>
        /// The left margin in inches.
        /// </summary>
        public double LeftMargin
        {
            get { return fLeftMargin; }
            set
            {
                fLeftMargin = value;
                if (fLeftMargin < 0) fLeftMargin = 0;
                HasPageMargins = true;
            }
        }

        internal double fRightMargin;
        /// <summary>
        /// The right margin in inches.
        /// </summary>
        public double RightMargin
        {
            get { return fRightMargin; }
            set
            {
                fRightMargin = value;
                if (fRightMargin < 0) fRightMargin = 0;
                HasPageMargins = true;
            }
        }

        internal double fTopMargin;
        /// <summary>
        /// The top margin in inches.
        /// </summary>
        public double TopMargin
        {
            get { return fTopMargin; }
            set
            {
                fTopMargin = value;
                if (fTopMargin < 0) fTopMargin = 0;
                HasPageMargins = true;
            }
        }

        internal double fBottomMargin;
        /// <summary>
        /// The bottom margin in inches.
        /// </summary>
        public double BottomMargin
        {
            get { return fBottomMargin; }
            set
            {
                fBottomMargin = value;
                if (fBottomMargin < 0) fBottomMargin = 0;
                HasPageMargins = true;
            }
        }

        internal double fHeaderMargin;
        /// <summary>
        /// The header margin in inches.
        /// </summary>
        public double HeaderMargin
        {
            get { return fHeaderMargin; }
            set
            {
                fHeaderMargin = value;
                if (fHeaderMargin < 0) fHeaderMargin = 0;
                HasPageMargins = true;
            }
        }

        internal double fFooterMargin;
        /// <summary>
        /// The footer margin in inches.
        /// </summary>
        public double FooterMargin
        {
            get { return fFooterMargin; }
            set
            {
                if (fFooterMargin < 0) fFooterMargin = 0;
                HasPageMargins = true;
            }
        }

        internal bool HasPageSetup
        {
            get
            {
                return this.PaperSize != SLPaperSizeValues.LetterPaper || this.FirstPageNumber != 1
                    || this.Scale != 100 || this.FitToWidth != 1 || this.FitToHeight != 1
                    || this.PageOrder != PageOrderValues.DownThenOver || this.Orientation != OrientationValues.Default
                    || !this.UsePrinterDefaults
                    || this.BlackAndWhite || this.Draft || this.CellComments != CellCommentsValues.None
                    || this.Errors != PrintErrorValues.Displayed || this.HorizontalDpi != 600
                    || this.VerticalDpi != 600 || this.Copies != 1;
            }
        }

        internal bool HasChartSheetPageSetup
        {
            get
            {
                return this.PaperSize != SLPaperSizeValues.LetterPaper || this.FirstPageNumber != 1
                    || this.Orientation != OrientationValues.Default
                    || !this.UsePrinterDefaults
                    || this.BlackAndWhite || this.Draft
                    || this.HorizontalDpi != 600
                    || this.VerticalDpi != 600 || this.Copies != 1;
            }
        }

        /// <summary>
        /// The paper size. The default is Letter.
        /// </summary>
        public SLPaperSizeValues PaperSize { get; set; }

        internal uint iFirstPageNumber;
        /// <summary>
        /// The page number set for the first printed page.
        /// </summary>
        public uint FirstPageNumber
        {
            get { return iFirstPageNumber; }
            set
            {
                iFirstPageNumber = value;
                if (iFirstPageNumber < 1) iFirstPageNumber = 1;
            }
        }

        internal uint iScale;
        /// <summary>
        /// The printing scale. This is read-only. This doesn't apply to chart sheets.
        /// </summary>
        public uint Scale
        {
            get { return iScale; }
        }

        internal uint iFitToWidth;
        /// <summary>
        /// The number of horizontal pages to fit into a printed page. This is read-only. This doesn't apply to chart sheets.
        /// </summary>
        public uint FitToWidth
        {
            get { return iFitToWidth; }
        }

        internal uint iFitToHeight;
        /// <summary>
        /// The number of vertical pages to fit into a printed page. This is read-only. This doesn't apply to chart sheets.
        /// </summary>
        public uint FitToHeight
        {
            get { return iFitToHeight; }
        }

        /// <summary>
        /// Page order when printed. This doesn't apply to chart sheets.
        /// </summary>
        public PageOrderValues PageOrder { get; set; }

        /// <summary>
        /// Page orientation.
        /// </summary>
        public OrientationValues Orientation { get; set; }

        internal bool UsePrinterDefaults { get; set; }

        /// <summary>
        /// Specifies if the page is printed in black and white.
        /// </summary>
        public bool BlackAndWhite { get; set; }

        /// <summary>
        /// Specifies if the page is printed in draft mode (without graphics).
        /// </summary>
        public bool Draft { get; set; }

        /// <summary>
        /// Specifies how to print cell comments. This doesn't apply to chart sheets.
        /// </summary>
        public CellCommentsValues CellComments { get; set; }

        /// <summary>
        /// Specifies how to print for cells with errors. This doesn't apply to chart sheets.
        /// </summary>
        public PrintErrorValues Errors { get; set; }

        /// <summary>
        /// Horizontal print resolution.
        /// </summary>
        public uint HorizontalDpi { get; set; }

        /// <summary>
        /// Vertical print resolution.
        /// </summary>
        public uint VerticalDpi { get; set; }

        internal uint iCopies;
        /// <summary>
        /// The number of copies to print. The minimum number is 1 copy. There are no maximum number of copies, however Excel uses 9999 copies as a maximum.
        /// </summary>
        public uint Copies
        {
            get { return iCopies; }
            set
            {
                iCopies = value;
                if (iCopies < 1) iCopies = 1;
            }
        }

        internal bool HasHeaderFooter
        {
            get
            {
                return OddHeaderText.Length > 0 || OddFooterText.Length > 0 || EvenHeaderText.Length > 0
                    || EvenFooterText.Length > 0 || FirstHeaderText.Length > 0 || FirstFooterText.Length > 0
                    || DifferentOddEvenPages || DifferentFirstPage || !ScaleWithDocument || !AlignWithMargins;
            }
        }

        /// <summary>
        /// The text in the odd-numbered page header. Note that this is the default used.
        /// </summary>
        public string OddHeaderText { get; set; }

        /// <summary>
        /// The text in the odd-numbered page footer. Note that this is the default used.
        /// </summary>
        public string OddFooterText { get; set; }

        /// <summary>
        /// The text in the even-numbered page header. Note that this only activates when <see cref="DifferentOddEvenPages"/> is true.
        /// </summary>
        public string EvenHeaderText { get; set; }

        /// <summary>
        /// The text in the even-numbered page footer. Note that this only activates when <see cref="DifferentOddEvenPages"/> is true.
        /// </summary>
        public string EvenFooterText { get; set; }

        /// <summary>
        /// The text in the first page's header. Note that this only activates when <see cref="DifferentFirstPage"/> is true.
        /// </summary>
        public string FirstHeaderText { get; set; }

        /// <summary>
        /// The text in the first page's footer. Note that this only activates when <see cref="DifferentFirstPage"/> is true.
        /// </summary>
        public string FirstFooterText { get; set; }

        /// <summary>
        /// Specifies if different headers and footers are set for odd- and even-numbered pages.
        /// If false, then the text in odd-numbered page header and footer is used, even if there's text
        /// set in even-numbered page header and footer.
        /// </summary>
        public bool DifferentOddEvenPages { get; set; }

        /// <summary>
        /// Specifies if a different header and footer is set for the first page.
        /// If false, any text set in the first page header and footer is ignored.
        /// </summary>
        public bool DifferentFirstPage { get; set; }

        /// <summary>
        /// Scale with the document.
        /// </summary>
        public bool ScaleWithDocument { get; set; }

        /// <summary>
        /// Align header and footer margins with page margins.
        /// </summary>
        public bool AlignWithMargins { get; set; }

        private bool StrikeSwitch { get; set; }
        private bool SuperscriptSwitch { get; set; }
        private bool SubscriptSwitch { get; set; }
        private bool UnderlineSwitch { get; set; }
        private bool DoubleUnderlineSwitch { get; set; }
        private bool FontSizeSwitch { get; set; }
        private bool FontColorSwitch { get; set; }
        private bool FontStyleSwitch { get; set; }
        private SLHeaderFooterSection HFSection { get; set; }

        /// <summary>
        /// Initializes an instance of SLPageSettings. It is recommended to use GetPageSettings() of the SLDocument class.
        /// </summary>
        public SLPageSettings()
        {
            this.Initialize(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
        }

        internal SLPageSettings(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
            this.Initialize(ThemeColors, IndexedColors);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
        {
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

            this.SetAllNull();
            this.ResetSwitches();
        }

        private void SetAllNull()
        {
            this.SheetProperties = new SLSheetProperties(this.listThemeColors, this.listIndexedColors);

            this.bShowFormulas = null;
            this.bShowGridLines = null;
            this.bShowRowColumnHeaders = null;
            this.bShowRuler = null;
            this.vView = null;
            this.iZoomScale = null;
            this.iZoomScaleNormal = null;
            this.iZoomScalePageLayoutView = null;

            this.PrintHorizontalCentered = false;
            this.PrintVerticalCentered = false;
            this.PrintHeadings = false;
            this.PrintGridLines = false;
            this.PrintGridLinesSet = true;

            this.SetNormalMargins();
            this.HasPageMargins = false;

            this.PaperSize = SLPaperSizeValues.LetterPaper;
            this.FirstPageNumber = 1;
            this.iScale = 100;
            this.iFitToWidth = 1;
            this.iFitToHeight = 1;
            this.PageOrder = PageOrderValues.DownThenOver;
            this.Orientation = OrientationValues.Default;
            this.UsePrinterDefaults = true;
            this.BlackAndWhite = false;
            this.Draft = false;
            this.CellComments = CellCommentsValues.None;
            this.Errors = PrintErrorValues.Displayed;
            this.HorizontalDpi = 600;
            this.VerticalDpi = 600;
            this.Copies = 1;

            this.OddHeaderText = string.Empty;
            this.OddFooterText = string.Empty;
            this.EvenHeaderText = string.Empty;
            this.EvenFooterText = string.Empty;
            this.FirstHeaderText = string.Empty;
            this.FirstFooterText = string.Empty;
            this.DifferentOddEvenPages = false;
            this.DifferentFirstPage = false;
            this.ScaleWithDocument = true;
            this.AlignWithMargins = true;
        }

        private void ResetSwitches()
        {
            this.StrikeSwitch = false;
            this.SuperscriptSwitch = false;
            this.SubscriptSwitch = false;
            this.UnderlineSwitch = false;
            this.DoubleUnderlineSwitch = false;
            this.FontSizeSwitch = false;
            this.FontColorSwitch = false;
            this.FontStyleSwitch = false;
            this.HFSection = SLHeaderFooterSection.None;
        }

        /// <summary>
        /// Sets the tab color of the sheet.
        /// </summary>
        /// <param name="TabColor">The theme color to be used.</param>
        public void SetTabColor(SLThemeColorIndexValues TabColor)
        {
            this.SheetProperties.clrTabColor.SetThemeColor(TabColor);
            this.SheetProperties.HasTabColor = (this.SheetProperties.clrTabColor.Color.IsEmpty) ? false : true;
        }

        /// <summary>
        /// Sets the tab color of the sheet.
        /// </summary>
        /// <param name="TabColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetTabColor(SLThemeColorIndexValues TabColor, double Tint)
        {
            this.SheetProperties.clrTabColor.SetThemeColor(TabColor, Tint);
            this.SheetProperties.HasTabColor = (this.SheetProperties.clrTabColor.Color.IsEmpty) ? false : true;
        }

        /// <summary>
        /// Set normal margins.
        /// </summary>
        public void SetNormalMargins()
        {
            this.fTopMargin = SLConstants.NormalTopMargin;
            this.fBottomMargin = SLConstants.NormalBottomMargin;
            this.fLeftMargin = SLConstants.NormalLeftMargin;
            this.fRightMargin = SLConstants.NormalRightMargin;
            this.fHeaderMargin = SLConstants.NormalHeaderMargin;
            this.fFooterMargin = SLConstants.NormalFooterMargin;
            this.HasPageMargins = true;
        }

        /// <summary>
        /// Set wide margins.
        /// </summary>
        public void SetWideMargins()
        {
            this.fTopMargin = SLConstants.WideTopMargin;
            this.fBottomMargin = SLConstants.WideBottomMargin;
            this.fLeftMargin = SLConstants.WideLeftMargin;
            this.fRightMargin = SLConstants.WideRightMargin;
            this.fHeaderMargin = SLConstants.WideHeaderMargin;
            this.fFooterMargin = SLConstants.WideFooterMargin;
            this.HasPageMargins = true;
        }

        /// <summary>
        /// Set narrow margins.
        /// </summary>
        public void SetNarrowMargins()
        {
            this.fTopMargin = SLConstants.NarrowTopMargin;
            this.fBottomMargin = SLConstants.NarrowBottomMargin;
            this.fLeftMargin = SLConstants.NarrowLeftMargin;
            this.fRightMargin = SLConstants.NarrowRightMargin;
            this.fHeaderMargin = SLConstants.NarrowHeaderMargin;
            this.fFooterMargin = SLConstants.NarrowFooterMargin;
            this.HasPageMargins = true;
        }

        /// <summary>
        /// Adjust the page a given percentage of the normal size.
        /// </summary>
        /// <param name="ScalePercentage">The scale percentage between 10% and 400%.</param>
        public void ScalePage(uint ScalePercentage)
        {
            if (ScalePercentage < 10) ScalePercentage = 10;
            if (ScalePercentage > 400) ScalePercentage = 400;
            this.iScale = ScalePercentage;

            this.iFitToWidth = 1;
            this.iFitToHeight = 1;

            this.SheetProperties.FitToPage = false;
        }

        /// <summary>
        /// Fit to a given number of pages wide, and a given number of pages high.
        /// </summary>
        /// <param name="FitToWidth">Number of pages wide. Minimum is 1 page (default).</param>
        /// <param name="FitToHeight">Number of pages high. Minimum is 1 page (default).</param>
        public void ScalePage(uint FitToWidth, uint FitToHeight)
        {
            if (FitToWidth < 1) FitToWidth = 1;
            if (FitToHeight < 1) FitToHeight = 1;
            this.iFitToWidth = FitToWidth;
            this.iFitToHeight = FitToHeight;

            this.iScale = 100;

            this.SheetProperties.FitToPage = true;
        }

        // the switches are universal!
        // To do it "properly", we could have 6 versions of the switch variables
        // (1 for each type of odd/even/first header/footer).
        // Is this important? The "workaround" is to do each header/footer type
        // all the way through before working on another type. (which is more natural...)
        // The "bug" will appear if you append some styled text on say OddHeader,
        // then append some styled text on FirstFooter.
        // The switches are still assumed to work on OddHeader, but should be reset for
        // FirstFooter. This is fine until you go back to appending some styled text for
        // OddHeader.

        /// <summary>
        /// Append text to the odd-numbered page header.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendOddHeader(string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.OddHeader) ResetSwitches();
            this.OddHeaderText += Text;
            this.HFSection = SLHeaderFooterSection.OddHeader;
        }

        /// <summary>
        /// Append text to the odd-numbered page header.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendOddHeader(SLFont FontStyle, string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.OddHeader) ResetSwitches();
            this.OddHeaderText += string.Format("{0} {1}", this.StyleToAppend(FontStyle), Text);
            this.HFSection = SLHeaderFooterSection.OddHeader;
        }

        /// <summary>
        /// Append a format code to the odd-numbered page header.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendOddHeader(SLHeaderFooterFormatCodeValues Code)
        {
            if (this.HFSection != SLHeaderFooterSection.OddHeader) ResetSwitches();
            this.OddHeaderText += this.TextToAppend(Code);
            this.HFSection = SLHeaderFooterSection.OddHeader;
        }

        /// <summary>
        /// Append text to the odd-numbered page footer.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendOddFooter(string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.OddFooter) ResetSwitches();
            this.OddFooterText += Text;
            this.HFSection = SLHeaderFooterSection.OddFooter;
        }

        /// <summary>
        /// Append text to the odd-numbered page footer.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendOddFooter(SLFont FontStyle, string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.OddFooter) ResetSwitches();
            this.OddFooterText += string.Format("{0} {1}", this.StyleToAppend(FontStyle), Text);
            this.HFSection = SLHeaderFooterSection.OddFooter;
        }

        /// <summary>
        /// Append a format code to the odd-numbered page footer.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendOddFooter(SLHeaderFooterFormatCodeValues Code)
        {
            if (this.HFSection != SLHeaderFooterSection.OddFooter) ResetSwitches();
            this.OddFooterText += this.TextToAppend(Code);
            this.HFSection = SLHeaderFooterSection.OddFooter;
        }

        /// <summary>
        /// Append text to the even-numbered page header.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendEvenHeader(string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.EvenHeader) ResetSwitches();
            this.EvenHeaderText += Text;
            this.HFSection = SLHeaderFooterSection.EvenHeader;
        }

        /// <summary>
        /// Append text to the even-numbered page header.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendEvenHeader(SLFont FontStyle, string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.EvenHeader) ResetSwitches();
            this.EvenHeaderText += string.Format("{0} {1}", this.StyleToAppend(FontStyle), Text);
            this.HFSection = SLHeaderFooterSection.EvenHeader;
        }

        /// <summary>
        /// Append a format code to the even-numbered page header.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendEvenHeader(SLHeaderFooterFormatCodeValues Code)
        {
            if (this.HFSection != SLHeaderFooterSection.EvenHeader) ResetSwitches();
            this.EvenHeaderText += this.TextToAppend(Code);
            this.HFSection = SLHeaderFooterSection.EvenHeader;
        }

        /// <summary>
        /// Append text to the even-numbered page footer.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendEvenFooter(string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.EvenFooter) ResetSwitches();
            this.EvenFooterText += Text;
            this.HFSection = SLHeaderFooterSection.EvenFooter;
        }

        /// <summary>
        /// Append text to the even-numbered page footer.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendEvenFooter(SLFont FontStyle, string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.EvenFooter) ResetSwitches();
            this.EvenFooterText += string.Format("{0} {1}", this.StyleToAppend(FontStyle), Text);
            this.HFSection = SLHeaderFooterSection.EvenFooter;
        }

        /// <summary>
        /// Append a format code to the even-numbered page footer.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendEvenFooter(SLHeaderFooterFormatCodeValues Code)
        {
            if (this.HFSection != SLHeaderFooterSection.EvenFooter) ResetSwitches();
            this.EvenFooterText += this.TextToAppend(Code);
            this.HFSection = SLHeaderFooterSection.EvenFooter;
        }

        /// <summary>
        /// Append text to the first page header.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendFirstHeader(string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.FirstHeader) ResetSwitches();
            this.FirstHeaderText += Text;
            this.HFSection = SLHeaderFooterSection.FirstHeader;
        }

        /// <summary>
        /// Append text to the first page header.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendFirstHeader(SLFont FontStyle, string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.FirstHeader) ResetSwitches();
            this.FirstHeaderText += string.Format("{0} {1}", this.StyleToAppend(FontStyle), Text);
            this.HFSection = SLHeaderFooterSection.FirstHeader;
        }

        /// <summary>
        /// Append a format code to the first page header.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendFirstHeader(SLHeaderFooterFormatCodeValues Code)
        {
            if (this.HFSection != SLHeaderFooterSection.FirstHeader) ResetSwitches();
            this.FirstHeaderText += this.TextToAppend(Code);
            this.HFSection = SLHeaderFooterSection.FirstHeader;
        }

        /// <summary>
        /// Append text to the first page footer.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendFirstFooter(string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.FirstFooter) ResetSwitches();
            this.FirstFooterText += Text;
            this.HFSection = SLHeaderFooterSection.FirstFooter;
        }

        /// <summary>
        /// Append text to the first page footer.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendFirstFooter(SLFont FontStyle, string Text)
        {
            if (this.HFSection != SLHeaderFooterSection.FirstFooter) ResetSwitches();
            this.FirstFooterText += string.Format("{0} {1}", this.StyleToAppend(FontStyle), Text);
            this.HFSection = SLHeaderFooterSection.FirstFooter;
        }

        /// <summary>
        /// Append a format code to the first page footer.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendFirstFooter(SLHeaderFooterFormatCodeValues Code)
        {
            if (this.HFSection != SLHeaderFooterSection.FirstFooter) ResetSwitches();
            this.FirstFooterText += this.TextToAppend(Code);
            this.HFSection = SLHeaderFooterSection.FirstFooter;
        }

        /// <summary>
        /// Get the text from the left section of the header.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetLeftHeaderText()
        {
            return GetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Left);
        }

        /// <summary>
        /// Get the text from the left section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <returns>The text.</returns>
        public string GetLeftHeaderText(SLHeaderFooterTypeValues HeaderType)
        {
            return GetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Left);
        }

        /// <summary>
        /// Get the text from the center section of the header.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetCenterHeaderText()
        {
            return GetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Center);
        }

        /// <summary>
        /// Get the text from the center section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <returns>The text.</returns>
        public string GetCenterHeaderText(SLHeaderFooterTypeValues HeaderType)
        {
            return GetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Center);
        }

        /// <summary>
        /// Get the text from the right section of the header.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetRightHeaderText()
        {
            return GetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Right);
        }

        /// <summary>
        /// Get the text from the right section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <returns>The text.</returns>
        public string GetRightHeaderText(SLHeaderFooterTypeValues HeaderType)
        {
            return GetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Right);
        }

        /// <summary>
        /// Get the text from the left section of the footer.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetLeftFooterText()
        {
            return GetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Left);
        }

        /// <summary>
        /// Get the text from the left section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <returns>The text.</returns>
        public string GetLeftFooterText(SLHeaderFooterTypeValues FooterType)
        {
            return GetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Left);
        }

        /// <summary>
        /// Get the text from the center section of the footer.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetCenterFooterText()
        {
            return GetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Center);
        }

        /// <summary>
        /// Get the text from the center section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <returns>The text.</returns>
        public string GetCenterFooterText(SLHeaderFooterTypeValues FooterType)
        {
            return GetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Center);
        }

        /// <summary>
        /// Get the text from the right section of the footer.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetRightFooterText()
        {
            return GetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Right);
        }

        /// <summary>
        /// Get the text from the right section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <returns>The text.</returns>
        public string GetRightFooterText(SLHeaderFooterTypeValues FooterType)
        {
            return GetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Right);
        }

        /// <summary>
        /// Set the text of the left section of the header.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetLeftHeaderText(string Text)
        {
            SetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Left, Text);
        }

        /// <summary>
        /// Set the text of the left section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <param name="Text">The text.</param>
        public void SetLeftHeaderText(SLHeaderFooterTypeValues HeaderType, string Text)
        {
            SetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Left, Text);
        }

        /// <summary>
        /// Set the text of the center section of the header.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetCenterHeaderText(string Text)
        {
            SetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Center, Text);
        }

        /// <summary>
        /// Set the text of the center section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <param name="Text">The text.</param>
        public void SetCenterHeaderText(SLHeaderFooterTypeValues HeaderType, string Text)
        {
            SetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Center, Text);
        }

        /// <summary>
        /// Set the text of the right section of the header.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetRightHeaderText(string Text)
        {
            SetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Right, Text);
        }

        /// <summary>
        /// Set the text of the right section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <param name="Text">The text.</param>
        public void SetRightHeaderText(SLHeaderFooterTypeValues HeaderType, string Text)
        {
            SetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Right, Text);
        }

        /// <summary>
        /// Set the text of the left section of the footer.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetLeftFooterText(string Text)
        {
            SetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Left, Text);
        }

        /// <summary>
        /// Set the text of the left section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <param name="Text">The text.</param>
        public void SetLeftFooterText(SLHeaderFooterTypeValues FooterType, string Text)
        {
            SetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Left, Text);
        }

        /// <summary>
        /// Set the text of the center section of the footer.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetCenterFooterText(string Text)
        {
            SetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Center, Text);
        }

        /// <summary>
        /// Set the text of the center section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <param name="Text">The text.</param>
        public void SetCenterFooterText(SLHeaderFooterTypeValues FooterType, string Text)
        {
            SetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Center, Text);
        }

        /// <summary>
        /// Set the text of the right section of the footer.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetRightFooterText(string Text)
        {
            SetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Right, Text);
        }

        /// <summary>
        /// Set the text of the right section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <param name="Text">The text.</param>
        public void SetRightFooterText(SLHeaderFooterTypeValues FooterType, string Text)
        {
            SetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Right, Text);
        }

        private string GetHeaderFooterText(bool IsHeader, SLHeaderFooterTypeValues Type, SLHeaderFooterSectionValues Section)
        {
            string result = string.Empty;
            string sLeft = string.Empty, sCenter = string.Empty, sRight = string.Empty;
            if (IsHeader)
            {
                if (Type == SLHeaderFooterTypeValues.Even) this.SplitHeaderFooterText(this.EvenHeaderText, out sLeft, out sCenter, out sRight);
                else if (Type == SLHeaderFooterTypeValues.First) this.SplitHeaderFooterText(this.FirstHeaderText, out sLeft, out sCenter, out sRight);
                else this.SplitHeaderFooterText(this.OddHeaderText, out sLeft, out sCenter, out sRight);
            }
            else
            {
                if (Type == SLHeaderFooterTypeValues.Even) this.SplitHeaderFooterText(this.EvenFooterText, out sLeft, out sCenter, out sRight);
                else if (Type == SLHeaderFooterTypeValues.First) this.SplitHeaderFooterText(this.FirstFooterText, out sLeft, out sCenter, out sRight);
                else this.SplitHeaderFooterText(this.OddFooterText, out sLeft, out sCenter, out sRight);
            }

            if (Section == SLHeaderFooterSectionValues.Left) result = sLeft;
            else if (Section == SLHeaderFooterSectionValues.Right) result = sRight;
            else result = sCenter;

            result = TranslateToUserFriendlyCode(result);

            return result;
        }

        private void SetHeaderFooterText(bool IsHeader, SLHeaderFooterTypeValues Type, SLHeaderFooterSectionValues Section, string Text)
        {
            string result = TranslateToInternalCode(Text);
            string sLeft = string.Empty, sCenter = string.Empty, sRight = string.Empty;
            if (IsHeader)
            {
                if (Type == SLHeaderFooterTypeValues.Even) this.SplitHeaderFooterText(this.EvenHeaderText, out sLeft, out sCenter, out sRight);
                else if (Type == SLHeaderFooterTypeValues.First) this.SplitHeaderFooterText(this.FirstHeaderText, out sLeft, out sCenter, out sRight);
                else this.SplitHeaderFooterText(this.OddHeaderText, out sLeft, out sCenter, out sRight);
            }
            else
            {
                if (Type == SLHeaderFooterTypeValues.Even) this.SplitHeaderFooterText(this.EvenFooterText, out sLeft, out sCenter, out sRight);
                else if (Type == SLHeaderFooterTypeValues.First) this.SplitHeaderFooterText(this.FirstFooterText, out sLeft, out sCenter, out sRight);
                else this.SplitHeaderFooterText(this.OddFooterText, out sLeft, out sCenter, out sRight);
            }

            if (Section == SLHeaderFooterSectionValues.Left) sLeft = result;
            else if (Section == SLHeaderFooterSectionValues.Right) sRight = result;
            else sCenter = result;

            result = string.Empty;

            if (sLeft.Length > 0) result += "&L" + sLeft;
            if (sCenter.Length > 0) result += "&C" + sCenter;
            if (sRight.Length > 0) result += "&R" + sRight;

            if (IsHeader)
            {
                if (Type == SLHeaderFooterTypeValues.Even) this.EvenHeaderText = result;
                else if (Type == SLHeaderFooterTypeValues.First) this.FirstHeaderText = result;
                else this.OddHeaderText = result;
            }
            else
            {
                if (Type == SLHeaderFooterTypeValues.Even) this.EvenFooterText = result;
                else if (Type == SLHeaderFooterTypeValues.First) this.FirstFooterText = result;
                else this.OddFooterText = result;
            }
        }

        private string TranslateToUserFriendlyCode(string HeaderFooterText)
        {
            string result = HeaderFooterText;
            result = Regex.Replace(result, "&[Pp]", "&[Page]");
            result = Regex.Replace(result, "&[Nn]", "&[Pages]");
            result = Regex.Replace(result, "&[Dd]", "&[Date]");
            result = Regex.Replace(result, "&[Tt]", "&[Time]");
            result = Regex.Replace(result, "&[Zz]", "&[Path]");
            result = Regex.Replace(result, "&[Ff]", "&[File]");
            result = Regex.Replace(result, "&[Aa]", "&[Tab]");

            return result;
        }

        private string TranslateToInternalCode(string HeaderFooterText)
        {
            string result = HeaderFooterText;
            result = Regex.Replace(result, "&\\[Page\\]", "&P");
            result = Regex.Replace(result, "&\\[Pages\\]", "&N");
            result = Regex.Replace(result, "&\\[Date\\]", "&D");
            result = Regex.Replace(result, "&\\[Time\\]", "&T");
            result = Regex.Replace(result, "&\\[Path\\]", "&Z");
            result = Regex.Replace(result, "&\\[File\\]", "&F");
            result = Regex.Replace(result, "&\\[Tab\\]", "&A");

            return result;
        }

        private string StyleToAppend(SLFont ft)
        {
            string result = string.Empty;

            string sBoldItalic = string.Empty;
            if ((ft.Bold != null && ft.Bold.Value) && (ft.Italic != null && ft.Italic.Value))
            {
                sBoldItalic = "Bold Italic";
            }
            else if (ft.Bold != null && ft.Bold.Value)
            {
                sBoldItalic = "Bold";
            }
            else if (ft.Italic != null && ft.Italic.Value)
            {
                sBoldItalic = "Italic";
            }
            else
            {
                sBoldItalic = "Regular";
            }

            // if it's bold or italic, there must at least be a font name or font scheme
            if (!sBoldItalic.Equals("Regular"))
            {
                if (ft.FontName == null || ft.FontName.Length == 0)
                {
                    if (ft.FontScheme == FontSchemeValues.None) ft.FontScheme = FontSchemeValues.Minor;
                }
            }

            string sFontStyle = string.Empty;
            if (this.FontStyleSwitch)
            {
                sFontStyle = this.FontStyleToAppend(ft, sBoldItalic);
                // must write something
                if (sFontStyle.Length == 0) sFontStyle = "&\"-,Regular\"";
            }
            else
            {
                sFontStyle = this.FontStyleToAppend(ft, sBoldItalic);
            }
            
            if (sFontStyle.Length == 0 || sFontStyle.Equals("&\"-,Regular\""))
            {
                this.FontStyleSwitch = false;
            }
            else
            {
                this.FontStyleSwitch = true;
            }

            result += sFontStyle;

            if (this.FontSizeSwitch)
            {
                // font size switch is on, so must write something
                if (ft.FontSize != null)
                {
                    result += string.Format("&{0}", (int)ft.FontSize);
                }
                else
                {
                    result += string.Format("&{0}", (int)SLConstants.DefaultFontSize);
                }
                this.FontSizeSwitch = false;
            }
            else
            {
                if (ft.FontSize != null && (int)ft.FontSize != (int)SLConstants.DefaultFontSize)
                {
                    result += string.Format("&{0}", (int)ft.FontSize);
                    this.FontSizeSwitch = true;
                }
            }

            if (this.StrikeSwitch)
            {
                // already in strikethrough mode
                // so only write something if given font style has no strikethrough
                if (ft.Strike == null || (ft.Strike != null && !ft.Strike.Value))
                {
                    result += "&S";
                    this.StrikeSwitch = false;
                }
            }
            else
            {
                if (ft.Strike != null && ft.Strike.Value)
                {
                    result += "&S";
                    this.StrikeSwitch = true;
                }
            }

            if (this.SuperscriptSwitch)
            {
                // already in superscript mode
                // so only write something if given font style has no superscript
                if (!ft.HasVerticalAlignment
                    || (ft.HasVerticalAlignment && ft.VerticalAlignment != VerticalAlignmentRunValues.Superscript))
                {
                    result += "&X";
                    this.SuperscriptSwitch = false;
                }
            }
            else
            {
                if (ft.HasVerticalAlignment && ft.VerticalAlignment == VerticalAlignmentRunValues.Superscript)
                {
                    result += "&X";
                    this.SuperscriptSwitch = true;
                }
            }

            if (this.SubscriptSwitch)
            {
                // already in subscript mode
                // so only write something if given font style has no subscript
                if (!ft.HasVerticalAlignment
                    || (ft.HasVerticalAlignment && ft.VerticalAlignment != VerticalAlignmentRunValues.Subscript))
                {
                    result += "&Y";
                    this.SubscriptSwitch = false;
                }
            }
            else
            {
                if (ft.HasVerticalAlignment && ft.VerticalAlignment == VerticalAlignmentRunValues.Subscript)
                {
                    result += "&Y";
                    this.SubscriptSwitch = true;
                }
            }

            if (this.UnderlineSwitch)
            {
                // already in underline mode
                // so only write something if given font style has no underline
                if (!ft.HasUnderline
                    || (ft.HasUnderline && ft.Underline != UnderlineValues.Single))
                {
                    // take care of SingleAccounting?
                    result += "&U";
                    this.UnderlineSwitch = false;
                }
            }
            else
            {
                if (ft.HasUnderline && ft.Underline == UnderlineValues.Single)
                {
                    // take care of SingleAccounting?
                    result += "&U";
                    this.UnderlineSwitch = true;
                }
            }

            if (this.DoubleUnderlineSwitch)
            {
                // already in double underline mode
                // so only write something if given font style has no double underline
                if (!ft.HasUnderline
                    || (ft.HasUnderline && ft.Underline == UnderlineValues.Double))
                {
                    // take care of DoubleAccounting?
                    result += "&E";
                    this.DoubleUnderlineSwitch = false;
                }
            }
            else
            {
                if (ft.HasUnderline && ft.Underline == UnderlineValues.Double)
                {
                    // take care of DoubleAccounting?
                    result += "&E";
                    this.DoubleUnderlineSwitch = true;
                }
            }

            if (this.FontColorSwitch)
            {
                if (ft.HasFontColor)
                {
                    result += this.FontColorToAppend(ft.clrFontColor);
                }
                else
                {
                    result += "&K01+000";
                    this.FontColorSwitch = false;
                }
            }
            else
            {
                if (ft.HasFontColor)
                {
                    result += this.FontColorToAppend(ft.clrFontColor);
                    this.FontColorSwitch = true;
                }
            }

            return result;
        }

        private string FontStyleToAppend(SLFont ft, string BoldItalic)
        {
            string result = string.Empty;

            if (ft.HasFontScheme)
            {
                if (ft.FontScheme == FontSchemeValues.Minor)
                {
                    result = string.Format("&\"-,{0}\"", BoldItalic);
                }
                else if (ft.FontScheme == FontSchemeValues.Major)
                {
                    result = string.Format("&\"+,{0}\"", BoldItalic);
                }
                else
                {
                    if (ft.FontName != null && ft.FontName.Length > 0)
                    {
                        result = string.Format("&\"{1},{0}\"", BoldItalic, ft.FontName);
                    }
                    else
                    {
                        result = string.Format("&\"-,{0}\"", BoldItalic);
                    }
                }
            }
            else if (ft.FontName != null && ft.FontName.Length > 0)
            {
                result = string.Format("&\"{1},{0}\"", BoldItalic, ft.FontName);
            }

            return result;
        }

        private string FontColorToAppend(SLColor clr)
        {
            string result = "&K01+000";

            if (clr.Theme != null)
            {
                double fTint = 0.0;
                bool bPositive = true;
                string sTint = string.Empty;
                if (clr.Tint != null)
                {
                    fTint = clr.Tint.Value;
                }

                if (fTint < 0)
                {
                    fTint = -fTint;
                    bPositive = false;
                }
                sTint = fTint.ToString(CultureInfo.InvariantCulture).Replace(".", "").PadRight(3, '0').Substring(0, 3);

                result = string.Format("&K{0}{1}{2}", clr.Theme.Value.ToString("d2"), bPositive ? "+" : "-", sTint);
            }
            else
            {
                result = string.Format("&K{0}{1}{2}", clr.Color.R.ToString("X2"), clr.Color.G.ToString("X2"), clr.Color.B.ToString("X2"));
            }

            return result;
        }

        private string TextToAppend(SLHeaderFooterFormatCodeValues Code)
        {
            string result = string.Empty;
            switch (Code)
            {
                case SLHeaderFooterFormatCodeValues.Left:
                    result = "&L";
                    ResetSwitches();
                    break;
                case SLHeaderFooterFormatCodeValues.Center:
                    result = "&C";
                    ResetSwitches();
                    break;
                case SLHeaderFooterFormatCodeValues.Right:
                    result = "&R";
                    ResetSwitches();
                    break;
                case SLHeaderFooterFormatCodeValues.PageNumber:
                    result = "&P";
                    break;
                case SLHeaderFooterFormatCodeValues.NumberOfPages:
                    result = "&N";
                    break;
                case SLHeaderFooterFormatCodeValues.Date:
                    result = "&D";
                    break;
                case SLHeaderFooterFormatCodeValues.Time:
                    result = "&T";
                    break;
                case SLHeaderFooterFormatCodeValues.FilePath:
                    result = "&Z";
                    break;
                case SLHeaderFooterFormatCodeValues.FileName:
                    result = "&F";
                    break;
                case SLHeaderFooterFormatCodeValues.SheetName:
                    result = "&A";
                    break;
                case SLHeaderFooterFormatCodeValues.ResetFont:
                    if (this.FontStyleSwitch) result += "&\"-,Regular\"";
                    if (this.FontSizeSwitch) result += string.Format("&{0}", (int)SLConstants.DefaultFontSize);
                    if (this.StrikeSwitch) result += "&S";
                    if (this.SuperscriptSwitch) result += "&X";
                    if (this.SubscriptSwitch) result += "&Y";
                    if (this.UnderlineSwitch) result += "&U";
                    if (this.DoubleUnderlineSwitch) result += "&E";
                    if (this.FontColorSwitch) result += "&K01+000";
                    ResetSwitches();
                    break;
            }
            return result;
        }

        private void SplitHeaderFooterText(string Text, out string Left, out string Center, out string Right)
        {
            Left = string.Empty;
            Center = string.Empty;
            Right = string.Empty;

            StringBuilder sbLeft = new StringBuilder();
            StringBuilder sbCenter = new StringBuilder();
            StringBuilder sbRight = new StringBuilder();

            // 0-left, 1-center, 2-right
            int iChoice = 1;

            for (int i = 0; i < Text.Length; ++i)
            {
                if (Text[i] == '&')
                {
                    if ((i + 1) < Text.Length)
                    {
                        // still within string length
                        if (Text[i + 1] == 'L' || Text[i + 1] == 'l')
                        {
                            iChoice = 0;
                            ++i;
                        }
                        else if (Text[i + 1] == 'C' || Text[i + 1] == 'c')
                        {
                            iChoice = 1;
                            ++i;
                        }
                        else if (Text[i + 1] == 'R' || Text[i + 1] == 'r')
                        {
                            iChoice = 2;
                            ++i;
                        }
                        else
                        {
                            // we're appending basically the ampersand
                            if (iChoice == 0) sbLeft.Append(Text[i]);
                            else if (iChoice == 2) sbRight.Append(Text[i]);
                            else sbCenter.Append(Text[i]);
                        }
                    }
                    else
                    {
                        // we're appending basically the ampersand
                        if (iChoice == 0) sbLeft.Append(Text[i]);
                        else if (iChoice == 2) sbRight.Append(Text[i]);
                        else sbCenter.Append(Text[i]);
                    }
                }
                else
                {
                    if (iChoice == 0) sbLeft.Append(Text[i]);
                    else if (iChoice == 2) sbRight.Append(Text[i]);
                    else sbCenter.Append(Text[i]);
                }
            }

            Left = sbLeft.ToString();
            Center = sbCenter.ToString();
            Right = sbRight.ToString();
        }

        internal void ImportPrintOptions(PrintOptions po)
        {
            if (po.HorizontalCentered != null) this.PrintHorizontalCentered = po.HorizontalCentered.Value;
            if (po.VerticalCentered != null) this.PrintVerticalCentered = po.VerticalCentered.Value;
            if (po.Headings != null) this.PrintHeadings = po.Headings.Value;
            if (po.GridLines != null) this.PrintGridLines = po.GridLines.Value;
            if (po.GridLinesSet != null) this.PrintGridLinesSet = po.GridLinesSet.Value;
        }

        internal PrintOptions ExportPrintOptions()
        {
            PrintOptions po = new PrintOptions();
            if (this.PrintHorizontalCentered) po.HorizontalCentered = true;
            if (this.PrintVerticalCentered) po.VerticalCentered = true;
            if (this.PrintHeadings) po.Headings = true;
            if (this.PrintGridLines) po.GridLines = true;
            if (!this.PrintGridLinesSet) po.GridLinesSet = false;

            return po;
        }

        internal void ImportPageMargins(PageMargins pm)
        {
            if (pm.Left != null) this.LeftMargin = pm.Left.Value;
            if (pm.Right != null) this.RightMargin = pm.Right.Value;
            if (pm.Top != null) this.TopMargin = pm.Top.Value;
            if (pm.Bottom != null) this.BottomMargin = pm.Bottom.Value;
            if (pm.Header != null) this.HeaderMargin = pm.Header.Value;
            if (pm.Footer != null) this.FooterMargin = pm.Footer.Value;
        }

        internal PageMargins ExportPageMargins()
        {
            PageMargins pm = new PageMargins();
            pm.Left = this.LeftMargin;
            pm.Right = this.RightMargin;
            pm.Top = this.TopMargin;
            pm.Bottom = this.BottomMargin;
            pm.Header = this.HeaderMargin;
            pm.Footer = this.FooterMargin;

            return pm;
        }

        internal void ImportPageSetup(PageSetup ps)
        {
            if (ps.PaperSize != null)
            {
                if (Enum.IsDefined(typeof(SLPaperSizeValues), ps.PaperSize.Value))
                {
                    this.PaperSize = (SLPaperSizeValues)ps.PaperSize.Value;
                }
                else
                {
                    this.PaperSize = SLPaperSizeValues.LetterPaper;
                }
            }

            if (ps.Scale != null) this.iScale = ps.Scale.Value;
            if (ps.FirstPageNumber != null) this.iFirstPageNumber = ps.FirstPageNumber.Value;
            if (ps.FitToWidth != null) this.iFitToWidth = ps.FitToWidth.Value;
            if (ps.FitToHeight != null) this.iFitToHeight = ps.FitToHeight.Value;
            if (ps.PageOrder != null) this.PageOrder = ps.PageOrder.Value;
            if (ps.Orientation != null) this.Orientation = ps.Orientation.Value;
            if (ps.UsePrinterDefaults != null) this.UsePrinterDefaults = ps.UsePrinterDefaults.Value;
            if (ps.BlackAndWhite != null) this.BlackAndWhite = ps.BlackAndWhite.Value;
            if (ps.Draft != null) this.Draft = ps.Draft.Value;
            if (ps.CellComments != null) this.CellComments = ps.CellComments.Value;
            if (ps.Errors != null) this.Errors = ps.Errors.Value;
            if (ps.HorizontalDpi != null) this.HorizontalDpi = ps.HorizontalDpi.Value;
            if (ps.VerticalDpi != null) this.VerticalDpi = ps.VerticalDpi.Value;
            if (ps.Copies != null) this.Copies = ps.Copies.Value;
        }

        internal PageSetup ExportPageSetup()
        {
            PageSetup ps = new PageSetup();
            if (this.PaperSize != SLPaperSizeValues.LetterPaper) ps.PaperSize = (uint)this.PaperSize;
            if (this.Scale != 100) ps.Scale = this.Scale;
            if (this.FitToWidth != 1 || this.FitToHeight != 1)
            {
                ps.FitToWidth = this.FitToWidth;
                ps.FitToHeight = this.FitToHeight;
            }
            if (this.FirstPageNumber != 1)
            {
                ps.FirstPageNumber = this.FirstPageNumber;
                ps.UseFirstPageNumber = true;
            }
            if (this.PageOrder != PageOrderValues.DownThenOver) ps.PageOrder = this.PageOrder;
            if (this.Orientation != OrientationValues.Default) ps.Orientation = this.Orientation;
            if (!this.UsePrinterDefaults) ps.UsePrinterDefaults = this.UsePrinterDefaults;
            if (this.BlackAndWhite) ps.BlackAndWhite = this.BlackAndWhite;
            if (this.Draft) ps.Draft = this.Draft;
            if (this.CellComments != CellCommentsValues.None) ps.CellComments = this.CellComments;
            if (this.Errors != PrintErrorValues.Displayed) ps.Errors = this.Errors;
            if (this.HorizontalDpi != 600) ps.HorizontalDpi = this.HorizontalDpi;
            if (this.VerticalDpi != 600) ps.VerticalDpi = this.VerticalDpi;
            if (this.Copies != 1) ps.Copies = this.Copies;

            return ps;
        }

        internal void ImportChartSheetPageSetup(ChartSheetPageSetup ps)
        {
            if (ps.PaperSize != null)
            {
                if (Enum.IsDefined(typeof(SLPaperSizeValues), ps.PaperSize.Value))
                {
                    this.PaperSize = (SLPaperSizeValues)ps.PaperSize.Value;
                }
                else
                {
                    this.PaperSize = SLPaperSizeValues.LetterPaper;
                }
            }

            if (ps.FirstPageNumber != null) this.iFirstPageNumber = ps.FirstPageNumber.Value;
            if (ps.Orientation != null) this.Orientation = ps.Orientation.Value;
            if (ps.UsePrinterDefaults != null) this.UsePrinterDefaults = ps.UsePrinterDefaults.Value;
            if (ps.BlackAndWhite != null) this.BlackAndWhite = ps.BlackAndWhite.Value;
            if (ps.Draft != null) this.Draft = ps.Draft.Value;
            if (ps.HorizontalDpi != null) this.HorizontalDpi = ps.HorizontalDpi.Value;
            if (ps.VerticalDpi != null) this.VerticalDpi = ps.VerticalDpi.Value;
            if (ps.Copies != null) this.Copies = ps.Copies.Value;
        }

        internal ChartSheetPageSetup ExportChartSheetPageSetup()
        {
            ChartSheetPageSetup ps = new ChartSheetPageSetup();
            if (this.PaperSize != SLPaperSizeValues.LetterPaper) ps.PaperSize = (uint)this.PaperSize;
            if (this.FirstPageNumber != 1)
            {
                ps.FirstPageNumber = this.FirstPageNumber;
                ps.UseFirstPageNumber = true;
            }
            if (this.Orientation != OrientationValues.Default) ps.Orientation = this.Orientation;
            if (!this.UsePrinterDefaults) ps.UsePrinterDefaults = this.UsePrinterDefaults;
            if (this.BlackAndWhite) ps.BlackAndWhite = this.BlackAndWhite;
            if (this.Draft) ps.Draft = this.Draft;
            if (this.HorizontalDpi != 600) ps.HorizontalDpi = this.HorizontalDpi;
            if (this.VerticalDpi != 600) ps.VerticalDpi = this.VerticalDpi;
            if (this.Copies != 1) ps.Copies = this.Copies;

            return ps;
        }

        internal void ImportHeaderFooter(HeaderFooter hf)
        {
            if (hf.OddHeader != null) this.OddHeaderText = hf.OddHeader.Text;
            if (hf.OddFooter != null) this.OddFooterText = hf.OddFooter.Text;
            if (hf.EvenHeader != null) this.EvenHeaderText = hf.EvenHeader.Text;
            if (hf.EvenFooter != null) this.EvenFooterText = hf.EvenFooter.Text;
            if (hf.FirstHeader != null) this.FirstHeaderText = hf.FirstHeader.Text;
            if (hf.FirstFooter != null) this.FirstFooterText = hf.FirstFooter.Text;
            if (hf.DifferentOddEven != null) this.DifferentOddEvenPages = hf.DifferentOddEven.Value;
            if (hf.DifferentFirst != null) this.DifferentFirstPage = hf.DifferentFirst.Value;
            if (hf.ScaleWithDoc != null) this.ScaleWithDocument = hf.ScaleWithDoc.Value;
            if (hf.AlignWithMargins != null) this.AlignWithMargins = hf.AlignWithMargins.Value;
        }

        internal HeaderFooter ExportHeaderFooter()
        {
            HeaderFooter hf = new HeaderFooter();
            if (this.OddHeaderText.Length > 0) hf.OddHeader = new OddHeader(this.OddHeaderText);
            if (this.OddFooterText.Length > 0) hf.OddFooter = new OddFooter(this.OddFooterText);
            if (this.EvenHeaderText.Length > 0) hf.EvenHeader = new EvenHeader(this.EvenHeaderText);
            if (this.EvenFooterText.Length > 0) hf.EvenFooter = new EvenFooter(this.EvenFooterText);
            if (this.FirstHeaderText.Length > 0) hf.FirstHeader = new FirstHeader(this.FirstHeaderText);
            if (this.FirstFooterText.Length > 0) hf.FirstFooter = new FirstFooter(this.FirstFooterText);
            if (this.DifferentOddEvenPages) hf.DifferentOddEven = this.DifferentOddEvenPages;
            if (this.DifferentFirstPage) hf.DifferentFirst = this.DifferentFirstPage;
            if (!this.ScaleWithDocument) hf.ScaleWithDoc = this.ScaleWithDocument;
            if (!this.AlignWithMargins) hf.AlignWithMargins = this.AlignWithMargins;

            return hf;
        }

        internal SLSheetView ExportSLSheetView()
        {
            SLSheetView sv = new SLSheetView();
            if (this.bShowFormulas != null) sv.ShowFormulas = this.bShowFormulas.Value;
            if (this.bShowGridLines != null) sv.ShowGridLines = this.bShowGridLines.Value;
            if (this.bShowRowColumnHeaders != null) sv.ShowRowColHeaders = this.bShowRowColumnHeaders.Value;
            if (this.bShowRuler != null) sv.ShowRuler = this.bShowRuler.Value;
            if (this.vView != null) sv.View = this.vView.Value;
            if (this.iZoomScale != null) sv.ZoomScale = this.iZoomScale.Value;
            if (this.iZoomScaleNormal != null) sv.ZoomScaleNormal = this.iZoomScaleNormal.Value;
            if (this.iZoomScalePageLayoutView != null) sv.ZoomScalePageLayoutView = this.iZoomScalePageLayoutView.Value;

            return sv;
        }

        internal SLPageSettings Clone()
        {
            SLPageSettings ps = new SLPageSettings(this.listThemeColors, this.listIndexedColors);

            ps.SheetProperties = this.SheetProperties.Clone();

            ps.bShowFormulas = this.bShowFormulas;
            ps.bShowGridLines = this.bShowGridLines;
            ps.bShowRowColumnHeaders = this.bShowRowColumnHeaders;
            ps.bShowRuler = this.bShowRuler;
            ps.vView = this.vView;
            ps.iZoomScale = this.iZoomScale;
            ps.iZoomScaleNormal = this.iZoomScaleNormal;
            ps.iZoomScalePageLayoutView = this.iZoomScalePageLayoutView;

            ps.PrintHorizontalCentered = this.PrintHorizontalCentered;
            ps.PrintVerticalCentered = this.PrintVerticalCentered;
            ps.PrintHeadings = this.PrintHeadings;
            ps.PrintGridLines = this.PrintGridLines;
            ps.PrintGridLinesSet = this.PrintGridLinesSet;

            ps.HasPageMargins = this.HasPageMargins;
            ps.fLeftMargin = this.fLeftMargin;
            ps.fRightMargin = this.fRightMargin;
            ps.fTopMargin = this.fTopMargin;
            ps.fBottomMargin = this.fBottomMargin;
            ps.fHeaderMargin = this.fHeaderMargin;
            ps.fFooterMargin = this.fFooterMargin;

            ps.PaperSize = this.PaperSize;
            ps.iFirstPageNumber = this.iFirstPageNumber;
            ps.iScale = this.iScale;
            ps.iFitToWidth = this.iFitToWidth;
            ps.iFitToHeight = this.iFitToHeight;
            ps.PageOrder = this.PageOrder;
            ps.Orientation = this.Orientation;
            ps.UsePrinterDefaults = this.UsePrinterDefaults;
            ps.BlackAndWhite = this.BlackAndWhite;
            ps.Draft = this.Draft;
            ps.CellComments = this.CellComments;
            ps.Errors = this.Errors;
            ps.HorizontalDpi = this.HorizontalDpi;
            ps.VerticalDpi = this.VerticalDpi;
            ps.iCopies = this.iCopies;

            ps.OddHeaderText = this.OddHeaderText;
            ps.OddFooterText = this.OddFooterText;
            ps.EvenHeaderText = this.EvenHeaderText;
            ps.EvenFooterText = this.EvenFooterText;
            ps.FirstHeaderText = this.FirstHeaderText;
            ps.FirstFooterText = this.FirstFooterText;
            ps.DifferentOddEvenPages = this.DifferentOddEvenPages;
            ps.DifferentFirstPage = this.DifferentFirstPage;
            ps.ScaleWithDocument = this.ScaleWithDocument;
            ps.AlignWithMargins = this.AlignWithMargins;

            return ps;
        }
    }
}
