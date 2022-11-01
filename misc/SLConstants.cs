using System;

namespace SpreadsheetLight
{
    internal class SLConstants
    {
        internal const string ApplicationName = "SpreadsheetLight";

        internal const string BlankSpreadsheetFileName = "Book1.xlsx";
        internal const string DefaultFirstSheetName = "Sheet1";
        internal const int RowLimit = 1048576;
        internal const int ColumnLimit = 16384;

        internal const string W3CDTF = "yyyy-MM-ddTHH:mm:ssZ";

        // The calculation ID is for the application (read: Excel) to determine which
        // version of the calculation engine is used. I surmise that the earlier the version,
        // the better it is. What if a later version (say Excel 2010) is used, but the resulting
        // spreadsheet is opened with an earlier version (say Excel 2007)? Does the earlier
        // version of the spreadsheet application know what to do with the later version and
        // possibly formulas that it knows nothing about?
        // But all this is moot because SpreadsheetLight doesn't have a calculation engine.
        // Frankly speaking this ID is specific to the application, and not really tied to Excel.
        // You could use a Mersenne prime or your birthday if you really want to.
        // But let's be honest, Excel is the most widely used spreadsheet application out there...
        // The default form is [version][build]
        // Excel 2007 version is 12. Don't know where 5725 came from.
        // We'll just use this (came from author's Office 2007 Small Business edition).
        internal const uint CalculationId = 125725;
        // This is from author's Office Home and Business 2010
        //internal const uint CalculationId = 145621;
        // This is from author's Office Home and Business 2013
        //internal const uint CalculationId = 152511;

        internal const string ErrorDivisionByZero = "#DIV/0!";
        internal const string ErrorNA = "#N/A";
        internal const string ErrorName = "#NAME?";
        internal const string ErrorNull = "#NULL!";
        internal const string ErrorNumber = "#NUM!";
        internal const string ErrorReference = "#REF!";
        internal const string ErrorValue = "#VALUE!";

        internal const int DegreeToAngleRepresentation = 60000;
        internal const long PointToEMU = 12700;
        internal const long InchToEMU = 914400;
        internal const long CentimeterToEMU = 360000;

        internal const int CustomNumberFormatIdStartIndex = 164;
        internal const string NumberFormatGeneral = "General";
        // this is for Excel 2007
        //internal const string DefaultTableStyle = "TableStyleMedium9";
        // this is for Excel 2010 and 2013
        internal const string DefaultTableStyle = "TableStyleMedium2";
        internal const string DefaultPivotStyle = "PivotStyleLight16";

        // it seems that as long as different numbers are used within a chart,
        // it is fine. And no I don't know how Excel gets those axis IDs...
        internal const int PrimaryAxis1 = 1;
        internal const int PrimaryAxis2 = 2;
        internal const int PrimaryAxis3 = 3;
        internal const int SecondaryAxis1 = 4;
        internal const int SecondaryAxis2 = 5;

        // values are intentionally weird to avoid conflict
        internal const string XmlStyleAttributeSeparator = "(S,}";
        internal const string XmlStyleElementAttributeSeparator = "(S|}";
        internal const string XmlAlignmentAttributeSeparator = "(A,}";
        internal const string XmlProtectionAttributeSeparator = "(P,}";
        internal const string XmlBorderAttributeSeparator = "(B,}";
        internal const string XmlBorderElementAttributeSeparator = "(B|}";
        internal const string XmlBorderPropertiesAttributeSeparator = "(BP,}";
        internal const string XmlBorderPropertiesElementAttributeSeparator = "(BP|}";
        internal const string XmlCellStyleAttributeSeparator = "(CS,}";
        internal const string XmlColorAttributeSeparator = "(CLR,}";
        internal const string XmlPatternFillAttributeSeparator = "(PFi,}";
        internal const string XmlPatternFillElementAttributeSeparator = "(PFi|}";
        internal const string XmlGradientFillAttributeSeparator = "(GFi,}";
        internal const string XmlGradientFillElementAttributeSeparator = "(GFi|}";
        internal const string XmlTableStyleAttributeSeparator = "(TS,}";
        internal const string XmlTableStyleElementAttributeSeparator = "(TS|}";

        // 120 DPI and 96 DPI have 21 and 17 pixels respectively as the width.
        // It seems that the filter image is a square, and has width/height equal
        // to 2 pixels less than the row height (1 pixel gap at the top and at the bottom).
        // I'm too lazy to do calculations of this sort, so I'm going to just use a fixed value.
        // I'm going to do a compromise and use 19 pixels. That should satisfy most cases, since
        // we're using it to do autofitting.
        internal const int AutoFilterPixelWidth = 19;
        internal const string AutoFitCacheSeparator = "(:,;}";

        //120 DPI, the little -/+ box is 19 pixels wide (including the preceding space) but the square box is 11 pixels by 11 pixels.
        //96 DPI, the corresponding values are 14 pixels and 9 pixels (by 9 pixels).
        // As a compromise, I'll just use a fixed value and be done with it.
        // This is used for autofitting columns. We can ignore for rows because the square box is too small.
        // What little -/+ box? If there are multiple row/column fields in a pivot table section, the box comes up.
        internal const int PivotTableFilterPixelWidth = 17;

        // 255 characters
        internal const double MaximumColumnWidth = 255;
        // 96 DPI is 409.5 pt (546 pixels), 120 DPI is 409.2 pt  (682 pixels)
        internal const double MaximumRowHeight = 409.2;

        // any string that won't be in danger of being *actually* formatted will do...
        internal const string GeneralFormatPlaceholder = "SLGENERAL";
        // specifically, no d,m,y,t,h,s but anything that can be date-formatted is out.
        internal const string ElapsedHourFormatPlaceholder = "SLQWR";
        internal const string ElapsedMinuteFormatPlaceholder = "SLWRQ";
        internal const string ElapsedSecondFormatPlaceholder = "SLRQW";

        internal const string PrintAreaDefinedName = "_xlnm.Print_Area";
        internal const string PrintTitlesDefinedName = "_xlnm.Print_Titles";
        internal const string CriteriaDefinedName = "_xlnm.Criteria";
        internal const string FilterDatabaseDefinedName = "_xlnm._FilterDatabase";
        internal const string ExtractDefinedName = "_xlnm.Extract";
        internal const string ConsolidateAreaDefinedName = "_xlnm.Consolidate_Area";
        internal const string DatabaseDefinedName = "_xlnm.Database";
        internal const string SheetTitleDefinedName = "_xlnm.Sheet_Title";

        internal const double NormalTopMargin = 0.75;
        internal const double NormalBottomMargin = 0.75;
        internal const double NormalLeftMargin = 0.7;
        internal const double NormalRightMargin = 0.7;
        internal const double NormalHeaderMargin = 0.3;
        internal const double NormalFooterMargin = 0.3;

        internal const double WideTopMargin = 1;
        internal const double WideBottomMargin = 1;
        internal const double WideLeftMargin = 1;
        internal const double WideRightMargin = 1;
        internal const double WideHeaderMargin = 0.5;
        internal const double WideFooterMargin = 0.5;

        internal const double NarrowTopMargin = 0.75;
        internal const double NarrowBottomMargin = 0.75;
        internal const double NarrowLeftMargin = 0.25;
        internal const double NarrowRightMargin = 0.25;
        internal const double NarrowHeaderMargin = 0.3;
        internal const double NarrowFooterMargin = 0.3;

        internal const double DefaultFontSize = 11;
        internal const double TitleFontSize = 18;
        internal const double Heading1FontSize = 15;
        internal const double Heading2FontSize = 13;

        // these were obtained from Excel for 120 DPI Calibri 11 pt
        internal const double DefaultCommentBoxWidth = 100.8;
        internal const double DefaultCommentBoxHeight = 60.6;

        internal const double DefaultCommentTopOffset = -0.42;
        internal const double DefaultCommentLeftOffset = 0.1725;

        // because I'm super upset that VML is so complicated...
        internal const int VmlTenMillionIterations = 10000000;

        internal const string NamespaceRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        internal const string NamespaceX = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal const string NamespaceX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        internal const string NamespaceX14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
        internal const string NamespaceXm = "http://schemas.microsoft.com/office/excel/2006/main";
        internal const string NamespaceMc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        internal const string NamespaceC = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        internal const string NamespaceC14 = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
        internal const string NamespaceA = "http://schemas.openxmlformats.org/drawingml/2006/main";
        internal const string NamespaceXdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

        internal const string SparklineExtensionUri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
        internal const string ConditionalFormattingExtensionUri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
        internal const string DataBarExtensionUri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";

        // Excel 2010 offers more image processing capabilities. However, there's a catch:
        // *We* have to do the actual image processing and create an image of the resulting
        // image processing. A .wdp (Windows Media Photo) is kept of the original image, and the
        // resulting image processed image will be saved as normal (like ImagePart).
        // I'm not ready to write image processing capabilities into SpreadsheetLight just yet...
        // At least not such that the resulting image looks like what Excel produces, which is the main
        // point. If it doesn't look like what Excel produced, then the user is likely to dismiss it
        // as "inferior" or "wrong".
        // So there's no real point in pursuing this. However, in case you (or when I'm insane enough)
        // want to take a look, you might find this useful:
        //http://msdn.microsoft.com/en-us/library/system.windows.media.imaging.wmpbitmapencoder.aspx
        // That's the encoder needed to create the .wdp file as the extension feed data.
        // Also the following Office/Excel extension URI's:
        //internal const string DrawingExtensionUri = "{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}";
        //internal const string DrawingLocalDpiExtensionUri = "{28A0092B-C50C-407E-A947-70E740481C1C}";

        // Office 2007 themes:
        // Office, Apex, Aspect, Civic, Concourse, Equity, Flow, Foundry, Median, Metro, Module,
        // Opulent, Oriel, Origin, Paper, Solstice, Technic, Trek, Urban, Verve

        // Office 2010 themes:
        // Office (same as 2007), Adjacency, Angles, Apothecary, Austin, BlackTie, Clarity, Composite,
        // Couture, Elemental, Essential, Executive, Grid, Hardcover, Horizon, Newsprint, Perspective,
        // Pushpin, Slipstream, Thatch, Waveform

        // Office.com themes:
        // Autumn, Decatur, Kilter, Macro, Mylar, Sketchbook, Soho, Spring, Summer, Thermal, Tradeshow
        // UrbanPop, Winter

        // Office 2013 themes (in order. The 2nd part seems to be from LiveContent (live.com?)):
        // Office (different from 2007 and 2010), Facet, Integral, Ion, Ion Boardroom, Organic, Retrospect,
        // Slice, Wisp, [start of next part], Banded, Basis, Celestial, Dividend, Frame, Mesh,
        // Metropolitan, Parallax, Quotable, Savon, View, Wood Type

        // Update: 26 July 2013. Moar themes in Excel 2013!
        // Berlin, Circuit, Damask, Depth, Droplet, Main Event, Slate, Vapor Trail

        internal const string OfficeThemeName = "SpreadsheetLight Office";
        internal const string OfficeThemeMajorLatinFont = "Cambria";
        internal const string OfficeThemeMinorLatinFont = "Calibri";
        internal const string OfficeThemeDark1Color = "000000";
        internal const string OfficeThemeLight1Color = "FFFFFF";
        internal const string OfficeThemeDark2Color = "1F497D";
        internal const string OfficeThemeLight2Color = "EEECE1";
        internal const string OfficeThemeAccent1Color = "4F81BD";
        internal const string OfficeThemeAccent2Color = "C0504D";
        internal const string OfficeThemeAccent3Color = "9BBB59";
        internal const string OfficeThemeAccent4Color = "8064A2";
        internal const string OfficeThemeAccent5Color = "4BACC6";
        internal const string OfficeThemeAccent6Color = "F79646";
        internal const string OfficeThemeHyperlink = "0000FF";
        internal const string OfficeThemeFollowedHyperlinkColor = "800080";

        internal const string Office2013ThemeName = "SpreadsheetLight Office2013";
        internal const string Office2013ThemeMajorLatinFont = "Calibri Light";
        internal const string Office2013ThemeMinorLatinFont = "Calibri";
        internal const string Office2013ThemeDark1Color = "000000";
        internal const string Office2013ThemeLight1Color = "FFFFFF";
        internal const string Office2013ThemeDark2Color = "44546A";
        internal const string Office2013ThemeLight2Color = "E7E6E6";
        internal const string Office2013ThemeAccent1Color = "5B9BD5";
        internal const string Office2013ThemeAccent2Color = "ED7D31";
        internal const string Office2013ThemeAccent3Color = "A5A5A5";
        internal const string Office2013ThemeAccent4Color = "FFC000";
        internal const string Office2013ThemeAccent5Color = "4472C4";
        internal const string Office2013ThemeAccent6Color = "70AD47";
        internal const string Office2013ThemeHyperlink = "0563C1";
        internal const string Office2013ThemeFollowedHyperlinkColor = "954F72";

        internal const string AdjacencyThemeName = "SpreadsheetLight Adjacency";
        internal const string AdjacencyThemeMajorLatinFont = "Cambria";
        internal const string AdjacencyThemeMinorLatinFont = "Calibri";
        internal const string AdjacencyThemeDark1Color = "2F2B20";
        internal const string AdjacencyThemeLight1Color = "FFFFFF";
        internal const string AdjacencyThemeDark2Color = "675E47";
        internal const string AdjacencyThemeLight2Color = "DFDCB7";
        internal const string AdjacencyThemeAccent1Color = "A9A57C";
        internal const string AdjacencyThemeAccent2Color = "9CBEBD";
        internal const string AdjacencyThemeAccent3Color = "D2CB6C";
        internal const string AdjacencyThemeAccent4Color = "95A39D";
        internal const string AdjacencyThemeAccent5Color = "C89F5D";
        internal const string AdjacencyThemeAccent6Color = "B1A089";
        internal const string AdjacencyThemeHyperlink = "D25814";
        internal const string AdjacencyThemeFollowedHyperlinkColor = "849A0A";

        internal const string AnglesThemeName = "SpreadsheetLight Angles";
        internal const string AnglesThemeMajorLatinFont = "Franklin Gothic Medium";
        internal const string AnglesThemeMinorLatinFont = "Franklin Gothic Book";
        internal const string AnglesThemeDark1Color = "000000";
        internal const string AnglesThemeLight1Color = "FFFFFF";
        internal const string AnglesThemeDark2Color = "434342";
        internal const string AnglesThemeLight2Color = "CDD7D9";
        internal const string AnglesThemeAccent1Color = "797B7E";
        internal const string AnglesThemeAccent2Color = "F96A1B";
        internal const string AnglesThemeAccent3Color = "08A1D9";
        internal const string AnglesThemeAccent4Color = "7C984A";
        internal const string AnglesThemeAccent5Color = "C2AD8D";
        internal const string AnglesThemeAccent6Color = "506E94";
        internal const string AnglesThemeHyperlink = "5F5F5F";
        internal const string AnglesThemeFollowedHyperlinkColor = "969696";

        internal const string ApexThemeName = "SpreadsheetLight Apex";
        internal const string ApexThemeMajorLatinFont = "Lucida Sans";
        internal const string ApexThemeMinorLatinFont = "Book Antiqua";
        internal const string ApexThemeDark1Color = "000000";
        internal const string ApexThemeLight1Color = "FFFFFF";
        internal const string ApexThemeDark2Color = "69676D";
        internal const string ApexThemeLight2Color = "C9C2D1";
        internal const string ApexThemeAccent1Color = "CEB966";
        internal const string ApexThemeAccent2Color = "9CB084";
        internal const string ApexThemeAccent3Color = "6BB1C9";
        internal const string ApexThemeAccent4Color = "6585CF";
        internal const string ApexThemeAccent5Color = "7E6BC9";
        internal const string ApexThemeAccent6Color = "A379BB";
        internal const string ApexThemeHyperlink = "410082";
        internal const string ApexThemeFollowedHyperlinkColor = "932968";

        internal const string ApothecaryThemeName = "SpreadsheetLight Apothecary";
        internal const string ApothecaryThemeMajorLatinFont = "Book Antiqua";
        internal const string ApothecaryThemeMinorLatinFont = "Century Gothic";
        internal const string ApothecaryThemeDark1Color = "000000";
        internal const string ApothecaryThemeLight1Color = "FFFFFF";
        internal const string ApothecaryThemeDark2Color = "564B3C";
        internal const string ApothecaryThemeLight2Color = "ECEDD1";
        internal const string ApothecaryThemeAccent1Color = "93A299";
        internal const string ApothecaryThemeAccent2Color = "CF543F";
        internal const string ApothecaryThemeAccent3Color = "B5AE53";
        internal const string ApothecaryThemeAccent4Color = "848058";
        internal const string ApothecaryThemeAccent5Color = "E8B54D";
        internal const string ApothecaryThemeAccent6Color = "786C71";
        internal const string ApothecaryThemeHyperlink = "CCCC00";
        internal const string ApothecaryThemeFollowedHyperlinkColor = "B2B2B2";

        internal const string AspectThemeName = "SpreadsheetLight Aspect";
        internal const string AspectThemeMajorLatinFont = "Verdana";
        internal const string AspectThemeMinorLatinFont = "Verdana";
        internal const string AspectThemeDark1Color = "000000";
        internal const string AspectThemeLight1Color = "FFFFFF";
        internal const string AspectThemeDark2Color = "323232";
        internal const string AspectThemeLight2Color = "E3DED1";
        internal const string AspectThemeAccent1Color = "F07F09";
        internal const string AspectThemeAccent2Color = "9F2936";
        internal const string AspectThemeAccent3Color = "1B587C";
        internal const string AspectThemeAccent4Color = "4E8542";
        internal const string AspectThemeAccent5Color = "604878";
        internal const string AspectThemeAccent6Color = "C19859";
        internal const string AspectThemeHyperlink = "6B9F25";
        internal const string AspectThemeFollowedHyperlinkColor = "B26B02";

        internal const string AustinThemeName = "SpreadsheetLight Austin";
        internal const string AustinThemeMajorLatinFont = "Century Gothic";
        internal const string AustinThemeMinorLatinFont = "Century Gothic";
        internal const string AustinThemeDark1Color = "000000";
        internal const string AustinThemeLight1Color = "FFFFFF";
        internal const string AustinThemeDark2Color = "3E3D2D";
        internal const string AustinThemeLight2Color = "CAF278";
        internal const string AustinThemeAccent1Color = "94C600";
        internal const string AustinThemeAccent2Color = "71685A";
        internal const string AustinThemeAccent3Color = "FF6700";
        internal const string AustinThemeAccent4Color = "909465";
        internal const string AustinThemeAccent5Color = "956B43";
        internal const string AustinThemeAccent6Color = "FEA022";
        internal const string AustinThemeHyperlink = "E68200";
        internal const string AustinThemeFollowedHyperlinkColor = "FFA94A";

        internal const string BlackTieThemeName = "SpreadsheetLight BlackTie";
        internal const string BlackTieThemeMajorLatinFont = "Garamond";
        internal const string BlackTieThemeMinorLatinFont = "Garamond";
        internal const string BlackTieThemeDark1Color = "000000";
        internal const string BlackTieThemeLight1Color = "FFFFFF";
        internal const string BlackTieThemeDark2Color = "46464A";
        internal const string BlackTieThemeLight2Color = "E3DCCF";
        internal const string BlackTieThemeAccent1Color = "6F6F74";
        internal const string BlackTieThemeAccent2Color = "A7B789";
        internal const string BlackTieThemeAccent3Color = "BEAE98";
        internal const string BlackTieThemeAccent4Color = "92A9B9";
        internal const string BlackTieThemeAccent5Color = "9C8265";
        internal const string BlackTieThemeAccent6Color = "8D6974";
        internal const string BlackTieThemeHyperlink = "67AABF";
        internal const string BlackTieThemeFollowedHyperlinkColor = "B1B5AB";

        internal const string CivicThemeName = "SpreadsheetLight Civic";
        internal const string CivicThemeMajorLatinFont = "Georgia";
        internal const string CivicThemeMinorLatinFont = "Georgia";
        internal const string CivicThemeDark1Color = "000000";
        internal const string CivicThemeLight1Color = "FFFFFF";
        internal const string CivicThemeDark2Color = "646B86";
        internal const string CivicThemeLight2Color = "C5D1D7";
        internal const string CivicThemeAccent1Color = "D16349";
        internal const string CivicThemeAccent2Color = "CCB400";
        internal const string CivicThemeAccent3Color = "8CADAE";
        internal const string CivicThemeAccent4Color = "8C7B70";
        internal const string CivicThemeAccent5Color = "8FB08C";
        internal const string CivicThemeAccent6Color = "D19049";
        internal const string CivicThemeHyperlink = "00A3D6";
        internal const string CivicThemeFollowedHyperlinkColor = "694F07";

        internal const string ClarityThemeName = "SpreadsheetLight Clarity";
        internal const string ClarityThemeMajorLatinFont = "Arial";
        internal const string ClarityThemeMinorLatinFont = "Arial";
        internal const string ClarityThemeDark1Color = "292934";
        internal const string ClarityThemeLight1Color = "FFFFFF";
        internal const string ClarityThemeDark2Color = "D2533C";
        internal const string ClarityThemeLight2Color = "F3F2DC";
        internal const string ClarityThemeAccent1Color = "93A299";
        internal const string ClarityThemeAccent2Color = "AD8F67";
        internal const string ClarityThemeAccent3Color = "726056";
        internal const string ClarityThemeAccent4Color = "4C5A6A";
        internal const string ClarityThemeAccent5Color = "808DA0";
        internal const string ClarityThemeAccent6Color = "79463D";
        internal const string ClarityThemeHyperlink = "0000FF";
        internal const string ClarityThemeFollowedHyperlinkColor = "800080";

        internal const string CompositeThemeName = "SpreadsheetLight Composite";
        internal const string CompositeThemeMajorLatinFont = "Calibri";
        internal const string CompositeThemeMinorLatinFont = "Calibri";
        internal const string CompositeThemeDark1Color = "000000";
        internal const string CompositeThemeLight1Color = "FFFFFF";
        internal const string CompositeThemeDark2Color = "5B6973";
        internal const string CompositeThemeLight2Color = "E7ECED";
        internal const string CompositeThemeAccent1Color = "98C723";
        internal const string CompositeThemeAccent2Color = "59B0B9";
        internal const string CompositeThemeAccent3Color = "DEAE00";
        internal const string CompositeThemeAccent4Color = "B77BB4";
        internal const string CompositeThemeAccent5Color = "E0773C";
        internal const string CompositeThemeAccent6Color = "A98D63";
        internal const string CompositeThemeHyperlink = "26CBEC";
        internal const string CompositeThemeFollowedHyperlinkColor = "598C8C";

        internal const string ConcourseThemeName = "SpreadsheetLight Concourse";
        internal const string ConcourseThemeMajorLatinFont = "Lucida Sans Unicode";
        internal const string ConcourseThemeMinorLatinFont = "Lucida Sans Unicode";
        internal const string ConcourseThemeDark1Color = "000000";
        internal const string ConcourseThemeLight1Color = "FFFFFF";
        internal const string ConcourseThemeDark2Color = "464646";
        internal const string ConcourseThemeLight2Color = "DEF5FA";
        internal const string ConcourseThemeAccent1Color = "2DA2BF";
        internal const string ConcourseThemeAccent2Color = "DA1F28";
        internal const string ConcourseThemeAccent3Color = "EB641B";
        internal const string ConcourseThemeAccent4Color = "39639D";
        internal const string ConcourseThemeAccent5Color = "474B78";
        internal const string ConcourseThemeAccent6Color = "7D3C4A";
        internal const string ConcourseThemeHyperlink = "FF8119";
        internal const string ConcourseThemeFollowedHyperlinkColor = "44B9E8";

        internal const string CoutureThemeName = "SpreadsheetLight Couture";
        internal const string CoutureThemeMajorLatinFont = "Garamond";
        internal const string CoutureThemeMinorLatinFont = "Garamond";
        internal const string CoutureThemeDark1Color = "000000";
        internal const string CoutureThemeLight1Color = "FFFFFF";
        internal const string CoutureThemeDark2Color = "37302A";
        internal const string CoutureThemeLight2Color = "D0CCB9";
        internal const string CoutureThemeAccent1Color = "9E8E5C";
        internal const string CoutureThemeAccent2Color = "A09781";
        internal const string CoutureThemeAccent3Color = "85776D";
        internal const string CoutureThemeAccent4Color = "AEAFA9";
        internal const string CoutureThemeAccent5Color = "8D878B";
        internal const string CoutureThemeAccent6Color = "6B6149";
        internal const string CoutureThemeHyperlink = "B6A272";
        internal const string CoutureThemeFollowedHyperlinkColor = "8A784F";

        internal const string ElementalThemeName = "SpreadsheetLight Elemental";
        internal const string ElementalThemeMajorLatinFont = "Palatino Linotype";
        internal const string ElementalThemeMinorLatinFont = "Palatino Linotype";
        internal const string ElementalThemeDark1Color = "000000";
        internal const string ElementalThemeLight1Color = "FFFFFF";
        internal const string ElementalThemeDark2Color = "242852";
        internal const string ElementalThemeLight2Color = "ACCBF9";
        internal const string ElementalThemeAccent1Color = "629DD1";
        internal const string ElementalThemeAccent2Color = "297FD5";
        internal const string ElementalThemeAccent3Color = "7F8FA9";
        internal const string ElementalThemeAccent4Color = "4A66AC";
        internal const string ElementalThemeAccent5Color = "5AA2AE";
        internal const string ElementalThemeAccent6Color = "9D90A0";
        internal const string ElementalThemeHyperlink = "9454C3";
        internal const string ElementalThemeFollowedHyperlinkColor = "3EBBF0";

        internal const string EquityThemeName = "SpreadsheetLight Equity";
        internal const string EquityThemeMajorLatinFont = "Franklin Gothic Book";
        internal const string EquityThemeMinorLatinFont = "Perpetua";
        internal const string EquityThemeDark1Color = "000000";
        internal const string EquityThemeLight1Color = "FFFFFF";
        internal const string EquityThemeDark2Color = "696464";
        internal const string EquityThemeLight2Color = "E9E5DC";
        internal const string EquityThemeAccent1Color = "D34817";
        internal const string EquityThemeAccent2Color = "9B2D1F";
        internal const string EquityThemeAccent3Color = "A28E6A";
        internal const string EquityThemeAccent4Color = "956251";
        internal const string EquityThemeAccent5Color = "918485";
        internal const string EquityThemeAccent6Color = "855D5D";
        internal const string EquityThemeHyperlink = "CC9900";
        internal const string EquityThemeFollowedHyperlinkColor = "96A9A9";

        internal const string EssentialThemeName = "SpreadsheetLight Essential";
        internal const string EssentialThemeMajorLatinFont = "Arial Black";
        internal const string EssentialThemeMinorLatinFont = "Arial";
        internal const string EssentialThemeDark1Color = "000000";
        internal const string EssentialThemeLight1Color = "FFFFFF";
        internal const string EssentialThemeDark2Color = "D1282E";
        internal const string EssentialThemeLight2Color = "C8C8B1";
        internal const string EssentialThemeAccent1Color = "7A7A7A";
        internal const string EssentialThemeAccent2Color = "F5C201";
        internal const string EssentialThemeAccent3Color = "526DB0";
        internal const string EssentialThemeAccent4Color = "989AAC";
        internal const string EssentialThemeAccent5Color = "DC5924";
        internal const string EssentialThemeAccent6Color = "B4B392";
        internal const string EssentialThemeHyperlink = "CC9900";
        internal const string EssentialThemeFollowedHyperlinkColor = "969696";

        internal const string ExecutiveThemeName = "SpreadsheetLight Executive";
        internal const string ExecutiveThemeMajorLatinFont = "Century Gothic";
        internal const string ExecutiveThemeMinorLatinFont = "Palatino Linotype";
        internal const string ExecutiveThemeDark1Color = "000000";
        internal const string ExecutiveThemeLight1Color = "FFFFFF";
        internal const string ExecutiveThemeDark2Color = "2F5897";
        internal const string ExecutiveThemeLight2Color = "E4E9EF";
        internal const string ExecutiveThemeAccent1Color = "6076B4";
        internal const string ExecutiveThemeAccent2Color = "9C5252";
        internal const string ExecutiveThemeAccent3Color = "E68422";
        internal const string ExecutiveThemeAccent4Color = "846648";
        internal const string ExecutiveThemeAccent5Color = "63891F";
        internal const string ExecutiveThemeAccent6Color = "758085";
        internal const string ExecutiveThemeHyperlink = "3399FF";
        internal const string ExecutiveThemeFollowedHyperlinkColor = "B2B2B2";

        internal const string FacetThemeName = "SpreadsheetLight Facet";
        internal const string FacetThemeMajorLatinFont = "Trebuchet MS";
        internal const string FacetThemeMinorLatinFont = "Trebuchet MS";
        internal const string FacetThemeDark1Color = "000000";
        internal const string FacetThemeLight1Color = "FFFFFF";
        internal const string FacetThemeDark2Color = "2C3C43";
        internal const string FacetThemeLight2Color = "EBEBEB";
        internal const string FacetThemeAccent1Color = "90C226";
        internal const string FacetThemeAccent2Color = "54A021";
        internal const string FacetThemeAccent3Color = "E6B91E";
        internal const string FacetThemeAccent4Color = "E76618";
        internal const string FacetThemeAccent5Color = "C42F1A";
        internal const string FacetThemeAccent6Color = "918655";
        internal const string FacetThemeHyperlink = "99CA3C";
        internal const string FacetThemeFollowedHyperlinkColor = "B9D181";

        internal const string FlowThemeName = "SpreadsheetLight Flow";
        internal const string FlowThemeMajorLatinFont = "Calibri";
        internal const string FlowThemeMinorLatinFont = "Constantia";
        internal const string FlowThemeDark1Color = "000000";
        internal const string FlowThemeLight1Color = "FFFFFF";
        internal const string FlowThemeDark2Color = "04617B";
        internal const string FlowThemeLight2Color = "DBF5F9";
        internal const string FlowThemeAccent1Color = "0F6FC6";
        internal const string FlowThemeAccent2Color = "009DD9";
        internal const string FlowThemeAccent3Color = "0BD0D9";
        internal const string FlowThemeAccent4Color = "10CF9B";
        internal const string FlowThemeAccent5Color = "7CCA62";
        internal const string FlowThemeAccent6Color = "A5C249";
        internal const string FlowThemeHyperlink = "E2D700";
        internal const string FlowThemeFollowedHyperlinkColor = "85DFD0";

        internal const string FoundryThemeName = "SpreadsheetLight Foundry";
        internal const string FoundryThemeMajorLatinFont = "Rockwell";
        internal const string FoundryThemeMinorLatinFont = "Rockwell";
        internal const string FoundryThemeDark1Color = "000000";
        internal const string FoundryThemeLight1Color = "FFFFFF";
        internal const string FoundryThemeDark2Color = "676A55";
        internal const string FoundryThemeLight2Color = "EAEBDE";
        internal const string FoundryThemeAccent1Color = "72A376";
        internal const string FoundryThemeAccent2Color = "B0CCB0";
        internal const string FoundryThemeAccent3Color = "A8CDD7";
        internal const string FoundryThemeAccent4Color = "C0BEAF";
        internal const string FoundryThemeAccent5Color = "CEC597";
        internal const string FoundryThemeAccent6Color = "E8B7B7";
        internal const string FoundryThemeHyperlink = "DB5353";
        internal const string FoundryThemeFollowedHyperlinkColor = "903638";

        internal const string GridThemeName = "SpreadsheetLight Grid";
        internal const string GridThemeMajorLatinFont = "Franklin Gothic Medium";
        internal const string GridThemeMinorLatinFont = "Franklin Gothic Medium";
        internal const string GridThemeDark1Color = "000000";
        internal const string GridThemeLight1Color = "FFFFFF";
        internal const string GridThemeDark2Color = "534949";
        internal const string GridThemeLight2Color = "CCD1B9";
        internal const string GridThemeAccent1Color = "C66951";
        internal const string GridThemeAccent2Color = "BF974D";
        internal const string GridThemeAccent3Color = "928B70";
        internal const string GridThemeAccent4Color = "87706B";
        internal const string GridThemeAccent5Color = "94734E";
        internal const string GridThemeAccent6Color = "6F777D";
        internal const string GridThemeHyperlink = "CC9900";
        internal const string GridThemeFollowedHyperlinkColor = "C0C0C0";

        internal const string HardcoverThemeName = "SpreadsheetLight Hardcover";
        internal const string HardcoverThemeMajorLatinFont = "Book Antiqua";
        internal const string HardcoverThemeMinorLatinFont = "Book Antiqua";
        internal const string HardcoverThemeDark1Color = "000000";
        internal const string HardcoverThemeLight1Color = "FFFFFF";
        internal const string HardcoverThemeDark2Color = "895D1D";
        internal const string HardcoverThemeLight2Color = "ECE9C6";
        internal const string HardcoverThemeAccent1Color = "873624";
        internal const string HardcoverThemeAccent2Color = "D6862D";
        internal const string HardcoverThemeAccent3Color = "D0BE40";
        internal const string HardcoverThemeAccent4Color = "877F6C";
        internal const string HardcoverThemeAccent5Color = "972109";
        internal const string HardcoverThemeAccent6Color = "AEB795";
        internal const string HardcoverThemeHyperlink = "CC9900";
        internal const string HardcoverThemeFollowedHyperlinkColor = "B2B2B2";

        internal const string HorizonThemeName = "SpreadsheetLight Horizon";
        internal const string HorizonThemeMajorLatinFont = "Arial Narrow";
        internal const string HorizonThemeMinorLatinFont = "Arial Narrow";
        internal const string HorizonThemeDark1Color = "000000";
        internal const string HorizonThemeLight1Color = "FFFFFF";
        internal const string HorizonThemeDark2Color = "1F2123";
        internal const string HorizonThemeLight2Color = "DC9E1F";
        internal const string HorizonThemeAccent1Color = "7E97AD";
        internal const string HorizonThemeAccent2Color = "CC8E60";
        internal const string HorizonThemeAccent3Color = "7A6A60";
        internal const string HorizonThemeAccent4Color = "B4936D";
        internal const string HorizonThemeAccent5Color = "67787B";
        internal const string HorizonThemeAccent6Color = "9D936F";
        internal const string HorizonThemeHyperlink = "646464";
        internal const string HorizonThemeFollowedHyperlinkColor = "969696";

        internal const string IntegralThemeName = "SpreadsheetLight Integral";
        internal const string IntegralThemeMajorLatinFont = "Tw Cen MT Condensed";
        internal const string IntegralThemeMinorLatinFont = "Tw Cen MT";
        internal const string IntegralThemeDark1Color = "000000";
        internal const string IntegralThemeLight1Color = "FFFFFF";
        internal const string IntegralThemeDark2Color = "335B74";
        internal const string IntegralThemeLight2Color = "DFE3E5";
        internal const string IntegralThemeAccent1Color = "1CADE4";
        internal const string IntegralThemeAccent2Color = "2683C6";
        internal const string IntegralThemeAccent3Color = "27CED7";
        internal const string IntegralThemeAccent4Color = "42BA97";
        internal const string IntegralThemeAccent5Color = "3E8853";
        internal const string IntegralThemeAccent6Color = "62A39F";
        internal const string IntegralThemeHyperlink = "6B9F25";
        internal const string IntegralThemeFollowedHyperlinkColor = "B26B02";

        internal const string IonThemeName = "SpreadsheetLight Ion";
        internal const string IonThemeMajorLatinFont = "Century Gothic";
        internal const string IonThemeMinorLatinFont = "Century Gothic";
        internal const string IonThemeDark1Color = "000000";
        internal const string IonThemeLight1Color = "FFFFFF";
        internal const string IonThemeDark2Color = "1E5155";
        internal const string IonThemeLight2Color = "EBEBEB";
        internal const string IonThemeAccent1Color = "B01513";
        internal const string IonThemeAccent2Color = "EA6312";
        internal const string IonThemeAccent3Color = "E6B729";
        internal const string IonThemeAccent4Color = "6AAC90";
        internal const string IonThemeAccent5Color = "54849A";
        internal const string IonThemeAccent6Color = "9E5E9B";
        internal const string IonThemeHyperlink = "58C1BA";
        internal const string IonThemeFollowedHyperlinkColor = "9DFFCB";

        internal const string IonBoardroomThemeName = "SpreadsheetLight Ion Boardroom";
        internal const string IonBoardroomThemeMajorLatinFont = "Century Gothic";
        internal const string IonBoardroomThemeMinorLatinFont = "Century Gothic";
        internal const string IonBoardroomThemeDark1Color = "000000";
        internal const string IonBoardroomThemeLight1Color = "FFFFFF";
        internal const string IonBoardroomThemeDark2Color = "3B3059";
        internal const string IonBoardroomThemeLight2Color = "EBEBEB";
        internal const string IonBoardroomThemeAccent1Color = "B31166";
        internal const string IonBoardroomThemeAccent2Color = "E33D6F";
        internal const string IonBoardroomThemeAccent3Color = "E45F3C";
        internal const string IonBoardroomThemeAccent4Color = "E9943A";
        internal const string IonBoardroomThemeAccent5Color = "9B6BF2";
        internal const string IonBoardroomThemeAccent6Color = "D53DD0";
        internal const string IonBoardroomThemeHyperlink = "8F8F8F";
        internal const string IonBoardroomThemeFollowedHyperlinkColor = "A5A5A5";

        internal const string MedianThemeName = "SpreadsheetLight Median";
        internal const string MedianThemeMajorLatinFont = "Tw Cen MT";
        internal const string MedianThemeMinorLatinFont = "Tw Cen MT";
        internal const string MedianThemeDark1Color = "000000";
        internal const string MedianThemeLight1Color = "FFFFFF";
        internal const string MedianThemeDark2Color = "775F55";
        internal const string MedianThemeLight2Color = "EBDDC3";
        internal const string MedianThemeAccent1Color = "94B6D2";
        internal const string MedianThemeAccent2Color = "DD8047";
        internal const string MedianThemeAccent3Color = "A5AB81";
        internal const string MedianThemeAccent4Color = "D8B25C";
        internal const string MedianThemeAccent5Color = "7BA79D";
        internal const string MedianThemeAccent6Color = "968C8C";
        internal const string MedianThemeHyperlink = "F7B615";
        internal const string MedianThemeFollowedHyperlinkColor = "704404";

        internal const string MetroThemeName = "SpreadsheetLight Metro";
        internal const string MetroThemeMajorLatinFont = "Consolas";
        internal const string MetroThemeMinorLatinFont = "Corbel";
        internal const string MetroThemeDark1Color = "000000";
        internal const string MetroThemeLight1Color = "FFFFFF";
        internal const string MetroThemeDark2Color = "4E5B6F";
        internal const string MetroThemeLight2Color = "D6ECFF";
        internal const string MetroThemeAccent1Color = "7FD13B";
        internal const string MetroThemeAccent2Color = "EA157A";
        internal const string MetroThemeAccent3Color = "FEB80A";
        internal const string MetroThemeAccent4Color = "00ADDC";
        internal const string MetroThemeAccent5Color = "738AC8";
        internal const string MetroThemeAccent6Color = "1AB39F";
        internal const string MetroThemeHyperlink = "EB8803";
        internal const string MetroThemeFollowedHyperlinkColor = "5F7791";

        internal const string ModuleThemeName = "SpreadsheetLight Module";
        internal const string ModuleThemeMajorLatinFont = "Corbel";
        internal const string ModuleThemeMinorLatinFont = "Corbel";
        internal const string ModuleThemeDark1Color = "000000";
        internal const string ModuleThemeLight1Color = "FFFFFF";
        internal const string ModuleThemeDark2Color = "5A6378";
        internal const string ModuleThemeLight2Color = "D4D4D6";
        internal const string ModuleThemeAccent1Color = "F0AD00";
        internal const string ModuleThemeAccent2Color = "60B5CC";
        internal const string ModuleThemeAccent3Color = "E66C7D";
        internal const string ModuleThemeAccent4Color = "6BB76D";
        internal const string ModuleThemeAccent5Color = "E88651";
        internal const string ModuleThemeAccent6Color = "C64847";
        internal const string ModuleThemeHyperlink = "168BBA";
        internal const string ModuleThemeFollowedHyperlinkColor = "680000";

        internal const string NewsprintThemeName = "SpreadsheetLight Newsprint";
        internal const string NewsprintThemeMajorLatinFont = "Impact";
        internal const string NewsprintThemeMinorLatinFont = "Times New Roman";
        internal const string NewsprintThemeDark1Color = "000000";
        internal const string NewsprintThemeLight1Color = "FFFFFF";
        internal const string NewsprintThemeDark2Color = "303030";
        internal const string NewsprintThemeLight2Color = "DEDEE0";
        internal const string NewsprintThemeAccent1Color = "AD0101";
        internal const string NewsprintThemeAccent2Color = "726056";
        internal const string NewsprintThemeAccent3Color = "AC956E";
        internal const string NewsprintThemeAccent4Color = "808DA9";
        internal const string NewsprintThemeAccent5Color = "424E5B";
        internal const string NewsprintThemeAccent6Color = "730E00";
        internal const string NewsprintThemeHyperlink = "D26900";
        internal const string NewsprintThemeFollowedHyperlinkColor = "D89243";

        internal const string OpulentThemeName = "SpreadsheetLight Opulent";
        internal const string OpulentThemeMajorLatinFont = "Trebuchet MS";
        internal const string OpulentThemeMinorLatinFont = "Trebuchet MS";
        internal const string OpulentThemeDark1Color = "000000";
        internal const string OpulentThemeLight1Color = "FFFFFF";
        internal const string OpulentThemeDark2Color = "B13F9A";
        internal const string OpulentThemeLight2Color = "F4E7ED";
        internal const string OpulentThemeAccent1Color = "B83D68";
        internal const string OpulentThemeAccent2Color = "AC66BB";
        internal const string OpulentThemeAccent3Color = "DE6C36";
        internal const string OpulentThemeAccent4Color = "F9B639";
        internal const string OpulentThemeAccent5Color = "CF6DA4";
        internal const string OpulentThemeAccent6Color = "FA8D3D";
        internal const string OpulentThemeHyperlink = "FFDE66";
        internal const string OpulentThemeFollowedHyperlinkColor = "D490C5";

        internal const string OrganicThemeName = "SpreadsheetLight Organic";
        internal const string OrganicThemeMajorLatinFont = "Garamond";
        internal const string OrganicThemeMinorLatinFont = "Garamond";
        internal const string OrganicThemeDark1Color = "000000";
        internal const string OrganicThemeLight1Color = "FFFFFF";
        internal const string OrganicThemeDark2Color = "212121";
        internal const string OrganicThemeLight2Color = "DADADA";
        internal const string OrganicThemeAccent1Color = "83992A";
        internal const string OrganicThemeAccent2Color = "3C9770";
        internal const string OrganicThemeAccent3Color = "44709D";
        internal const string OrganicThemeAccent4Color = "A23C33";
        internal const string OrganicThemeAccent5Color = "D97828";
        internal const string OrganicThemeAccent6Color = "DEB340";
        internal const string OrganicThemeHyperlink = "A8BF4D";
        internal const string OrganicThemeFollowedHyperlinkColor = "B4CA80";

        internal const string OrielThemeName = "SpreadsheetLight Oriel";
        internal const string OrielThemeMajorLatinFont = "Century Schoolbook";
        internal const string OrielThemeMinorLatinFont = "Century Schoolbook";
        internal const string OrielThemeDark1Color = "000000";
        internal const string OrielThemeLight1Color = "FFFFFF";
        internal const string OrielThemeDark2Color = "575F6D";
        internal const string OrielThemeLight2Color = "FFF39D";
        internal const string OrielThemeAccent1Color = "FE8637";
        internal const string OrielThemeAccent2Color = "7598D9";
        internal const string OrielThemeAccent3Color = "B32C16";
        internal const string OrielThemeAccent4Color = "F5CD2D";
        internal const string OrielThemeAccent5Color = "AEBAD5";
        internal const string OrielThemeAccent6Color = "777C84";
        internal const string OrielThemeHyperlink = "D2611C";
        internal const string OrielThemeFollowedHyperlinkColor = "3B435B";

        internal const string OriginThemeName = "SpreadsheetLight Origin";
        internal const string OriginThemeMajorLatinFont = "Bookman Old Style";
        internal const string OriginThemeMinorLatinFont = "Gill Sans MT";
        internal const string OriginThemeDark1Color = "000000";
        internal const string OriginThemeLight1Color = "FFFFFF";
        internal const string OriginThemeDark2Color = "464653";
        internal const string OriginThemeLight2Color = "DDE9EC";
        internal const string OriginThemeAccent1Color = "727CA3";
        internal const string OriginThemeAccent2Color = "9FB8CD";
        internal const string OriginThemeAccent3Color = "D2DA7A";
        internal const string OriginThemeAccent4Color = "FADA7A";
        internal const string OriginThemeAccent5Color = "B88472";
        internal const string OriginThemeAccent6Color = "8E736A";
        internal const string OriginThemeHyperlink = "B292CA";
        internal const string OriginThemeFollowedHyperlinkColor = "6B5680";

        internal const string PaperThemeName = "SpreadsheetLight Paper";
        internal const string PaperThemeMajorLatinFont = "Constantia";
        internal const string PaperThemeMinorLatinFont = "Constantia";
        internal const string PaperThemeDark1Color = "000000";
        internal const string PaperThemeLight1Color = "FFFFFF";
        internal const string PaperThemeDark2Color = "444D26";
        internal const string PaperThemeLight2Color = "FEFAC9";
        internal const string PaperThemeAccent1Color = "A5B592";
        internal const string PaperThemeAccent2Color = "F3A447";
        internal const string PaperThemeAccent3Color = "E7BC29";
        internal const string PaperThemeAccent4Color = "D092A7";
        internal const string PaperThemeAccent5Color = "9C85C0";
        internal const string PaperThemeAccent6Color = "809EC2";
        internal const string PaperThemeHyperlink = "8E58B6";
        internal const string PaperThemeFollowedHyperlinkColor = "7F6F6F";

        internal const string PerspectiveThemeName = "SpreadsheetLight Perspective";
        internal const string PerspectiveThemeMajorLatinFont = "Arial";
        internal const string PerspectiveThemeMinorLatinFont = "Arial";
        internal const string PerspectiveThemeDark1Color = "000000";
        internal const string PerspectiveThemeLight1Color = "FFFFFF";
        internal const string PerspectiveThemeDark2Color = "283138";
        internal const string PerspectiveThemeLight2Color = "FF8600";
        internal const string PerspectiveThemeAccent1Color = "838D9B";
        internal const string PerspectiveThemeAccent2Color = "D2610C";
        internal const string PerspectiveThemeAccent3Color = "80716A";
        internal const string PerspectiveThemeAccent4Color = "94147C";
        internal const string PerspectiveThemeAccent5Color = "5D5AD2";
        internal const string PerspectiveThemeAccent6Color = "6F6C7D";
        internal const string PerspectiveThemeHyperlink = "6187E3";
        internal const string PerspectiveThemeFollowedHyperlinkColor = "7B8EB8";

        internal const string PushpinThemeName = "SpreadsheetLight Pushpin";
        internal const string PushpinThemeMajorLatinFont = "Constantia";
        internal const string PushpinThemeMinorLatinFont = "Franklin Gothic Book";
        internal const string PushpinThemeDark1Color = "000000";
        internal const string PushpinThemeLight1Color = "FFFFFF";
        internal const string PushpinThemeDark2Color = "465E9C";
        internal const string PushpinThemeLight2Color = "CCDDEA";
        internal const string PushpinThemeAccent1Color = "FDA023";
        internal const string PushpinThemeAccent2Color = "AA2B1E";
        internal const string PushpinThemeAccent3Color = "71685C";
        internal const string PushpinThemeAccent4Color = "64A73B";
        internal const string PushpinThemeAccent5Color = "EB5605";
        internal const string PushpinThemeAccent6Color = "B9CA1A";
        internal const string PushpinThemeHyperlink = "D83E2C";
        internal const string PushpinThemeFollowedHyperlinkColor = "ED7D27";

        internal const string RetrospectThemeName = "SpreadsheetLight Retrospect";
        internal const string RetrospectThemeMajorLatinFont = "Calibri Light";
        internal const string RetrospectThemeMinorLatinFont = "Calibri";
        internal const string RetrospectThemeDark1Color = "000000";
        internal const string RetrospectThemeLight1Color = "FFFFFF";
        internal const string RetrospectThemeDark2Color = "637052";
        internal const string RetrospectThemeLight2Color = "CCDDEA";
        internal const string RetrospectThemeAccent1Color = "E48312";
        internal const string RetrospectThemeAccent2Color = "BD582C";
        internal const string RetrospectThemeAccent3Color = "865640";
        internal const string RetrospectThemeAccent4Color = "9B8357";
        internal const string RetrospectThemeAccent5Color = "C2BC80";
        internal const string RetrospectThemeAccent6Color = "94A088";
        internal const string RetrospectThemeHyperlink = "2998E3";
        internal const string RetrospectThemeFollowedHyperlinkColor = "8C8C8C";

        internal const string SliceThemeName = "SpreadsheetLight Slice";
        internal const string SliceThemeMajorLatinFont = "Century Gothic";
        internal const string SliceThemeMinorLatinFont = "Century Gothic";
        internal const string SliceThemeDark1Color = "000000";
        internal const string SliceThemeLight1Color = "FFFFFF";
        internal const string SliceThemeDark2Color = "146194";
        internal const string SliceThemeLight2Color = "76DBF4";
        internal const string SliceThemeAccent1Color = "052F61";
        internal const string SliceThemeAccent2Color = "A50E82";
        internal const string SliceThemeAccent3Color = "14967C";
        internal const string SliceThemeAccent4Color = "6A9E1F";
        internal const string SliceThemeAccent5Color = "E87D37";
        internal const string SliceThemeAccent6Color = "C62324";
        internal const string SliceThemeHyperlink = "0D2E46";
        internal const string SliceThemeFollowedHyperlinkColor = "356A95";

        internal const string SlipstreamThemeName = "SpreadsheetLight Slipstream";
        internal const string SlipstreamThemeMajorLatinFont = "Trebuchet MS";
        internal const string SlipstreamThemeMinorLatinFont = "Trebuchet MS";
        internal const string SlipstreamThemeDark1Color = "000000";
        internal const string SlipstreamThemeLight1Color = "FFFFFF";
        internal const string SlipstreamThemeDark2Color = "212745";
        internal const string SlipstreamThemeLight2Color = "B4DCFA";
        internal const string SlipstreamThemeAccent1Color = "4E67C8";
        internal const string SlipstreamThemeAccent2Color = "5ECCF3";
        internal const string SlipstreamThemeAccent3Color = "A7EA52";
        internal const string SlipstreamThemeAccent4Color = "5DCEAF";
        internal const string SlipstreamThemeAccent5Color = "FF8021";
        internal const string SlipstreamThemeAccent6Color = "F14124";
        internal const string SlipstreamThemeHyperlink = "56C7AA";
        internal const string SlipstreamThemeFollowedHyperlinkColor = "59A8D1";

        internal const string SolsticeThemeName = "SpreadsheetLight Solstice";
        internal const string SolsticeThemeMajorLatinFont = "Gill Sans MT";
        internal const string SolsticeThemeMinorLatinFont = "Gill Sans MT";
        internal const string SolsticeThemeDark1Color = "000000";
        internal const string SolsticeThemeLight1Color = "FFFFFF";
        internal const string SolsticeThemeDark2Color = "4F271C";
        internal const string SolsticeThemeLight2Color = "E7DEC9";
        internal const string SolsticeThemeAccent1Color = "3891A7";
        internal const string SolsticeThemeAccent2Color = "FEB80A";
        internal const string SolsticeThemeAccent3Color = "C32D2E";
        internal const string SolsticeThemeAccent4Color = "84AA33";
        internal const string SolsticeThemeAccent5Color = "964305";
        internal const string SolsticeThemeAccent6Color = "475A8D";
        internal const string SolsticeThemeHyperlink = "8DC765";
        internal const string SolsticeThemeFollowedHyperlinkColor = "AA8A14";

        internal const string TechnicThemeName = "SpreadsheetLight Technic";
        internal const string TechnicThemeMajorLatinFont = "Franklin Gothic Book";
        internal const string TechnicThemeMinorLatinFont = "Arial";
        internal const string TechnicThemeDark1Color = "000000";
        internal const string TechnicThemeLight1Color = "FFFFFF";
        internal const string TechnicThemeDark2Color = "3B3B3B";
        internal const string TechnicThemeLight2Color = "D4D2D0";
        internal const string TechnicThemeAccent1Color = "6EA0B0";
        internal const string TechnicThemeAccent2Color = "CCAF0A";
        internal const string TechnicThemeAccent3Color = "8D89A4";
        internal const string TechnicThemeAccent4Color = "748560";
        internal const string TechnicThemeAccent5Color = "9E9273";
        internal const string TechnicThemeAccent6Color = "7E848D";
        internal const string TechnicThemeHyperlink = "00C8C3";
        internal const string TechnicThemeFollowedHyperlinkColor = "A116E0";

        internal const string ThatchThemeName = "SpreadsheetLight Thatch";
        internal const string ThatchThemeMajorLatinFont = "Tw Cen MT";
        internal const string ThatchThemeMinorLatinFont = "Tw Cen MT";
        internal const string ThatchThemeDark1Color = "000000";
        internal const string ThatchThemeLight1Color = "FFFFFF";
        internal const string ThatchThemeDark2Color = "1D3641";
        internal const string ThatchThemeLight2Color = "DFE6D0";
        internal const string ThatchThemeAccent1Color = "759AA5";
        internal const string ThatchThemeAccent2Color = "CFC60D";
        internal const string ThatchThemeAccent3Color = "99987F";
        internal const string ThatchThemeAccent4Color = "90AC97";
        internal const string ThatchThemeAccent5Color = "FFAD1C";
        internal const string ThatchThemeAccent6Color = "B9AB6F";
        internal const string ThatchThemeHyperlink = "66AACD";
        internal const string ThatchThemeFollowedHyperlinkColor = "809DB3";

        internal const string TrekThemeName = "SpreadsheetLight Trek";
        internal const string TrekThemeMajorLatinFont = "Franklin Gothic Medium";
        internal const string TrekThemeMinorLatinFont = "Franklin Gothic Book";
        internal const string TrekThemeDark1Color = "000000";
        internal const string TrekThemeLight1Color = "FFFFFF";
        internal const string TrekThemeDark2Color = "4E3B30";
        internal const string TrekThemeLight2Color = "FBEEC9";
        internal const string TrekThemeAccent1Color = "F0A22E";
        internal const string TrekThemeAccent2Color = "A5644E";
        internal const string TrekThemeAccent3Color = "B58B80";
        internal const string TrekThemeAccent4Color = "C3986D";
        internal const string TrekThemeAccent5Color = "A19574";
        internal const string TrekThemeAccent6Color = "C17529";
        internal const string TrekThemeHyperlink = "AD1F1F";
        internal const string TrekThemeFollowedHyperlinkColor = "FFC42F";

        internal const string UrbanThemeName = "SpreadsheetLight Urban";
        internal const string UrbanThemeMajorLatinFont = "Trebuchet MS";
        internal const string UrbanThemeMinorLatinFont = "Georgia";
        internal const string UrbanThemeDark1Color = "000000";
        internal const string UrbanThemeLight1Color = "FFFFFF";
        internal const string UrbanThemeDark2Color = "424456";
        internal const string UrbanThemeLight2Color = "DEDEDE";
        internal const string UrbanThemeAccent1Color = "53548A";
        internal const string UrbanThemeAccent2Color = "438086";
        internal const string UrbanThemeAccent3Color = "A04DA3";
        internal const string UrbanThemeAccent4Color = "C4652D";
        internal const string UrbanThemeAccent5Color = "8B5D3D";
        internal const string UrbanThemeAccent6Color = "5C92B5";
        internal const string UrbanThemeHyperlink = "67AFBD";
        internal const string UrbanThemeFollowedHyperlinkColor = "C2A874";

        internal const string VerveThemeName = "SpreadsheetLight Verve";
        internal const string VerveThemeMajorLatinFont = "Century Gothic";
        internal const string VerveThemeMinorLatinFont = "Century Gothic";
        internal const string VerveThemeDark1Color = "000000";
        internal const string VerveThemeLight1Color = "FFFFFF";
        internal const string VerveThemeDark2Color = "666666";
        internal const string VerveThemeLight2Color = "D2D2D2";
        internal const string VerveThemeAccent1Color = "FF388C";
        internal const string VerveThemeAccent2Color = "E40059";
        internal const string VerveThemeAccent3Color = "9C007F";
        internal const string VerveThemeAccent4Color = "68007F";
        internal const string VerveThemeAccent5Color = "005BD3";
        internal const string VerveThemeAccent6Color = "00349E";
        internal const string VerveThemeHyperlink = "17BBFD";
        internal const string VerveThemeFollowedHyperlinkColor = "FF79C2";

        internal const string WaveformThemeName = "SpreadsheetLight Waveform";
        internal const string WaveformThemeMajorLatinFont = "Candara";
        internal const string WaveformThemeMinorLatinFont = "Candara";
        internal const string WaveformThemeDark1Color = "000000";
        internal const string WaveformThemeLight1Color = "FFFFFF";
        internal const string WaveformThemeDark2Color = "073E87";
        internal const string WaveformThemeLight2Color = "C6E7FC";
        internal const string WaveformThemeAccent1Color = "31B6FD";
        internal const string WaveformThemeAccent2Color = "4584D3";
        internal const string WaveformThemeAccent3Color = "5BD078";
        internal const string WaveformThemeAccent4Color = "A5D028";
        internal const string WaveformThemeAccent5Color = "F5C040";
        internal const string WaveformThemeAccent6Color = "05E0DB";
        internal const string WaveformThemeHyperlink = "0080FF";
        internal const string WaveformThemeFollowedHyperlinkColor = "5EAEFF";

        internal const string WispThemeName = "SpreadsheetLight Wisp";
        internal const string WispThemeMajorLatinFont = "Century Gothic";
        internal const string WispThemeMinorLatinFont = "Century Gothic";
        internal const string WispThemeDark1Color = "000000";
        internal const string WispThemeLight1Color = "FFFFFF";
        internal const string WispThemeDark2Color = "766F54";
        internal const string WispThemeLight2Color = "E3EACF";
        internal const string WispThemeAccent1Color = "A53010";
        internal const string WispThemeAccent2Color = "DE7E18";
        internal const string WispThemeAccent3Color = "9F8351";
        internal const string WispThemeAccent4Color = "728653";
        internal const string WispThemeAccent5Color = "92AA4C";
        internal const string WispThemeAccent6Color = "6AAC91";
        internal const string WispThemeHyperlink = "FB4A18";
        internal const string WispThemeFollowedHyperlinkColor = "FB9318";

        internal const string AutumnThemeName = "SpreadsheetLight Autumn";
        internal const string AutumnThemeMajorLatinFont = "Verdana";
        internal const string AutumnThemeMinorLatinFont = "Verdana";
        internal const string AutumnThemeDark1Color = "000000";
        internal const string AutumnThemeLight1Color = "FFFFFF";
        internal const string AutumnThemeDark2Color = "B01F0F";
        internal const string AutumnThemeLight2Color = "FF9000";
        internal const string AutumnThemeAccent1Color = "ED4600";
        internal const string AutumnThemeAccent2Color = "C4D73F";
        internal const string AutumnThemeAccent3Color = "FFCE2D";
        internal const string AutumnThemeAccent4Color = "FFA600";
        internal const string AutumnThemeAccent5Color = "ED5E00";
        internal const string AutumnThemeAccent6Color = "C62D03";
        internal const string AutumnThemeHyperlink = "408080";
        internal const string AutumnThemeFollowedHyperlinkColor = "5EAEAE";

        internal const string BandedThemeName = "SpreadsheetLight Banded";
        internal const string BandedThemeMajorLatinFont = "Corbel";
        internal const string BandedThemeMinorLatinFont = "Corbel";
        internal const string BandedThemeDark1Color = "2C2C2C";
        internal const string BandedThemeLight1Color = "FFFFFF";
        internal const string BandedThemeDark2Color = "099BDD";
        internal const string BandedThemeLight2Color = "F2F2F2";
        internal const string BandedThemeAccent1Color = "FFC000";
        internal const string BandedThemeAccent2Color = "A5D028";
        internal const string BandedThemeAccent3Color = "08CC78";
        internal const string BandedThemeAccent4Color = "F24099";
        internal const string BandedThemeAccent5Color = "828288";
        internal const string BandedThemeAccent6Color = "F56617";
        internal const string BandedThemeHyperlink = "005DBA";
        internal const string BandedThemeFollowedHyperlinkColor = "6C606A";

        internal const string BasisThemeName = "SpreadsheetLight Basis";
        internal const string BasisThemeMajorLatinFont = "Corbel";
        internal const string BasisThemeMinorLatinFont = "Corbel";
        internal const string BasisThemeDark1Color = "000000";
        internal const string BasisThemeLight1Color = "FFFFFF";
        internal const string BasisThemeDark2Color = "565349";
        internal const string BasisThemeLight2Color = "DDDDDD";
        internal const string BasisThemeAccent1Color = "A6B727";
        internal const string BasisThemeAccent2Color = "DF5327";
        internal const string BasisThemeAccent3Color = "FE9E00";
        internal const string BasisThemeAccent4Color = "418AB3";
        internal const string BasisThemeAccent5Color = "D7D447";
        internal const string BasisThemeAccent6Color = "818183";
        internal const string BasisThemeHyperlink = "F59E00";
        internal const string BasisThemeFollowedHyperlinkColor = "B2B2B2";

        internal const string BerlinThemeName = "SpreadsheetLight Berlin";
        internal const string BerlinThemeMajorLatinFont = "Trebuchet MS";
        internal const string BerlinThemeMinorLatinFont = "Trebuchet MS";
        internal const string BerlinThemeDark1Color = "000000";
        internal const string BerlinThemeLight1Color = "FFFFFF";
        internal const string BerlinThemeDark2Color = "9D360E";
        internal const string BerlinThemeLight2Color = "E7E6E6";
        internal const string BerlinThemeAccent1Color = "F09415";
        internal const string BerlinThemeAccent2Color = "C1B56B";
        internal const string BerlinThemeAccent3Color = "4BAF73";
        internal const string BerlinThemeAccent4Color = "5AA6C0";
        internal const string BerlinThemeAccent5Color = "D17DF9";
        internal const string BerlinThemeAccent6Color = "FA7E5C";
        internal const string BerlinThemeHyperlink = "FFAE3E";
        internal const string BerlinThemeFollowedHyperlinkColor = "FCC77E";

        internal const string CelestialThemeName = "SpreadsheetLight Celestial";
        internal const string CelestialThemeMajorLatinFont = "Calibri Light";
        internal const string CelestialThemeMinorLatinFont = "Calibri";
        internal const string CelestialThemeDark1Color = "000000";
        internal const string CelestialThemeLight1Color = "FFFFFF";
        internal const string CelestialThemeDark2Color = "18276C";
        internal const string CelestialThemeLight2Color = "EBEBEB";
        internal const string CelestialThemeAccent1Color = "AC3EC1";
        internal const string CelestialThemeAccent2Color = "477BD1";
        internal const string CelestialThemeAccent3Color = "46B298";
        internal const string CelestialThemeAccent4Color = "90BA4C";
        internal const string CelestialThemeAccent5Color = "DD9D31";
        internal const string CelestialThemeAccent6Color = "E25247";
        internal const string CelestialThemeHyperlink = "C573D2";
        internal const string CelestialThemeFollowedHyperlinkColor = "CCAEE8";

        internal const string CircuitThemeName = "SpreadsheetLight Circuit";
        internal const string CircuitThemeMajorLatinFont = "Tw Cen MT";
        internal const string CircuitThemeMinorLatinFont = "Tw Cen MT";
        internal const string CircuitThemeDark1Color = "000000";
        internal const string CircuitThemeLight1Color = "FFFFFF";
        internal const string CircuitThemeDark2Color = "134770";
        internal const string CircuitThemeLight2Color = "82FFFF";
        internal const string CircuitThemeAccent1Color = "9ACD4C";
        internal const string CircuitThemeAccent2Color = "FAA93A";
        internal const string CircuitThemeAccent3Color = "D35940";
        internal const string CircuitThemeAccent4Color = "B258D3";
        internal const string CircuitThemeAccent5Color = "63A0CC";
        internal const string CircuitThemeAccent6Color = "8AC4A7";
        internal const string CircuitThemeHyperlink = "B8FA56";
        internal const string CircuitThemeFollowedHyperlinkColor = "7AF8CC";

        internal const string DamaskThemeName = "SpreadsheetLight Damask";
        internal const string DamaskThemeMajorLatinFont = "Bookman Old Style";
        internal const string DamaskThemeMinorLatinFont = "Rockwell";
        internal const string DamaskThemeDark1Color = "000000";
        internal const string DamaskThemeLight1Color = "FFFFFF";
        internal const string DamaskThemeDark2Color = "2A5B7F";
        internal const string DamaskThemeLight2Color = "ABDAFC";
        internal const string DamaskThemeAccent1Color = "9EC544";
        internal const string DamaskThemeAccent2Color = "50BEA3";
        internal const string DamaskThemeAccent3Color = "4A9CCC";
        internal const string DamaskThemeAccent4Color = "9A66CA";
        internal const string DamaskThemeAccent5Color = "C54F71";
        internal const string DamaskThemeAccent6Color = "DE9C3C";
        internal const string DamaskThemeHyperlink = "6BA9DA";
        internal const string DamaskThemeFollowedHyperlinkColor = "A0BCD3";

        internal const string DecaturThemeName = "SpreadsheetLight Decatur";
        internal const string DecaturThemeMajorLatinFont = "Bodoni MT Condensed";
        internal const string DecaturThemeMinorLatinFont = "Franklin Gothic Book";
        internal const string DecaturThemeDark1Color = "000000";
        internal const string DecaturThemeLight1Color = "FFFFFF";
        internal const string DecaturThemeDark2Color = "55554A";
        internal const string DecaturThemeLight2Color = "D7DAE1";
        internal const string DecaturThemeAccent1Color = "F4680B";
        internal const string DecaturThemeAccent2Color = "ABB19F";
        internal const string DecaturThemeAccent3Color = "948774";
        internal const string DecaturThemeAccent4Color = "7EB8E7";
        internal const string DecaturThemeAccent5Color = "E3B651";
        internal const string DecaturThemeAccent6Color = "96756C";
        internal const string DecaturThemeHyperlink = "66AACD";
        internal const string DecaturThemeFollowedHyperlinkColor = "809DB3";

        internal const string DepthThemeName = "SpreadsheetLight Depth";
        internal const string DepthThemeMajorLatinFont = "Corbel";
        internal const string DepthThemeMinorLatinFont = "Corbel";
        internal const string DepthThemeDark1Color = "000000";
        internal const string DepthThemeLight1Color = "FFFFFF";
        internal const string DepthThemeDark2Color = "455F51";
        internal const string DepthThemeLight2Color = "94D7E4";
        internal const string DepthThemeAccent1Color = "41AEBD";
        internal const string DepthThemeAccent2Color = "97E9D5";
        internal const string DepthThemeAccent3Color = "A2CF49";
        internal const string DepthThemeAccent4Color = "608F3D";
        internal const string DepthThemeAccent5Color = "F4DE3A";
        internal const string DepthThemeAccent6Color = "FCB11C";
        internal const string DepthThemeHyperlink = "FBCA98";
        internal const string DepthThemeFollowedHyperlinkColor = "D3B86D";

        internal const string DividendThemeName = "SpreadsheetLight Dividend";
        internal const string DividendThemeMajorLatinFont = "Gill Sans MT";
        internal const string DividendThemeMinorLatinFont = "Gill Sans MT";
        internal const string DividendThemeDark1Color = "000000";
        internal const string DividendThemeLight1Color = "FFFFFF";
        internal const string DividendThemeDark2Color = "3D3D3D";
        internal const string DividendThemeLight2Color = "EBEBEB";
        internal const string DividendThemeAccent1Color = "4D1434";
        internal const string DividendThemeAccent2Color = "903163";
        internal const string DividendThemeAccent3Color = "B2324B";
        internal const string DividendThemeAccent4Color = "969FA7";
        internal const string DividendThemeAccent5Color = "66B1CE";
        internal const string DividendThemeAccent6Color = "40619D";
        internal const string DividendThemeHyperlink = "828282";
        internal const string DividendThemeFollowedHyperlinkColor = "A5A5A5";

        internal const string DropletThemeName = "SpreadsheetLight Droplet";
        internal const string DropletThemeMajorLatinFont = "Tw Cen MT";
        internal const string DropletThemeMinorLatinFont = "Tw Cen MT";
        internal const string DropletThemeDark1Color = "000000";
        internal const string DropletThemeLight1Color = "FFFFFF";
        internal const string DropletThemeDark2Color = "355071";
        internal const string DropletThemeLight2Color = "AABED7";
        internal const string DropletThemeAccent1Color = "2FA3EE";
        internal const string DropletThemeAccent2Color = "4BCAAD";
        internal const string DropletThemeAccent3Color = "86C157";
        internal const string DropletThemeAccent4Color = "D99C3F";
        internal const string DropletThemeAccent5Color = "CE6633";
        internal const string DropletThemeAccent6Color = "A35DD1";
        internal const string DropletThemeHyperlink = "56BCFE";
        internal const string DropletThemeFollowedHyperlinkColor = "97C5E3";

        internal const string FrameThemeName = "SpreadsheetLight Frame";
        internal const string FrameThemeMajorLatinFont = "Corbel";
        internal const string FrameThemeMinorLatinFont = "Corbel";
        internal const string FrameThemeDark1Color = "000000";
        internal const string FrameThemeLight1Color = "FFFFFF";
        internal const string FrameThemeDark2Color = "545454";
        internal const string FrameThemeLight2Color = "BFBFBF";
        internal const string FrameThemeAccent1Color = "40BAD2";
        internal const string FrameThemeAccent2Color = "FAB900";
        internal const string FrameThemeAccent3Color = "90BB23";
        internal const string FrameThemeAccent4Color = "EE7008";
        internal const string FrameThemeAccent5Color = "1AB39F";
        internal const string FrameThemeAccent6Color = "D5393D";
        internal const string FrameThemeHyperlink = "90BB23";
        internal const string FrameThemeFollowedHyperlinkColor = "EE7008";

        internal const string KilterThemeName = "SpreadsheetLight Kilter";
        internal const string KilterThemeMajorLatinFont = "Rockwell";
        internal const string KilterThemeMinorLatinFont = "Rockwell";
        internal const string KilterThemeDark1Color = "000000";
        internal const string KilterThemeLight1Color = "FFFFFF";
        internal const string KilterThemeDark2Color = "318FC5";
        internal const string KilterThemeLight2Color = "AEE8FB";
        internal const string KilterThemeAccent1Color = "76C5EF";
        internal const string KilterThemeAccent2Color = "FEA022";
        internal const string KilterThemeAccent3Color = "FF6700";
        internal const string KilterThemeAccent4Color = "70A525";
        internal const string KilterThemeAccent5Color = "A5D848";
        internal const string KilterThemeAccent6Color = "20768C";
        internal const string KilterThemeHyperlink = "7AB6E8";
        internal const string KilterThemeFollowedHyperlinkColor = "83B0D3";

        internal const string MacroThemeName = "SpreadsheetLight Macro";
        internal const string MacroThemeMajorLatinFont = "Calibri";
        internal const string MacroThemeMinorLatinFont = "Calibri";
        internal const string MacroThemeDark1Color = "000000";
        internal const string MacroThemeLight1Color = "FFFFFF";
        internal const string MacroThemeDark2Color = "3F3F4D";
        internal const string MacroThemeLight2Color = "DDDDDD";
        internal const string MacroThemeAccent1Color = "A51009";
        internal const string MacroThemeAccent2Color = "DE7014";
        internal const string MacroThemeAccent3Color = "704836";
        internal const string MacroThemeAccent4Color = "F2B431";
        internal const string MacroThemeAccent5Color = "7F221D";
        internal const string MacroThemeAccent6Color = "CDAC77";
        internal const string MacroThemeHyperlink = "F5B123";
        internal const string MacroThemeFollowedHyperlinkColor = "E19B0B";

        internal const string MainEventThemeName = "SpreadsheetLight Main Event";
        internal const string MainEventThemeMajorLatinFont = "Impact";
        internal const string MainEventThemeMinorLatinFont = "Impact";
        internal const string MainEventThemeDark1Color = "000000";
        internal const string MainEventThemeLight1Color = "FFFFFF";
        internal const string MainEventThemeDark2Color = "424242";
        internal const string MainEventThemeLight2Color = "C8C8C8";
        internal const string MainEventThemeAccent1Color = "B80E0F";
        internal const string MainEventThemeAccent2Color = "A6987D";
        internal const string MainEventThemeAccent3Color = "7F9A71";
        internal const string MainEventThemeAccent4Color = "64969F";
        internal const string MainEventThemeAccent5Color = "9B75B2";
        internal const string MainEventThemeAccent6Color = "80737A";
        internal const string MainEventThemeHyperlink = "F21213";
        internal const string MainEventThemeFollowedHyperlinkColor = "B6A394";

        internal const string MeshThemeName = "SpreadsheetLight Mesh";
        internal const string MeshThemeMajorLatinFont = "Century Gothic";
        internal const string MeshThemeMinorLatinFont = "Century Gothic";
        internal const string MeshThemeDark1Color = "000000";
        internal const string MeshThemeLight1Color = "FFFFFF";
        internal const string MeshThemeDark2Color = "363D46";
        internal const string MeshThemeLight2Color = "EBEBEB";
        internal const string MeshThemeAccent1Color = "6F6F6F";
        internal const string MeshThemeAccent2Color = "BFBFA5";
        internal const string MeshThemeAccent3Color = "DCD084";
        internal const string MeshThemeAccent4Color = "E7BF5F";
        internal const string MeshThemeAccent5Color = "E9A039";
        internal const string MeshThemeAccent6Color = "CF7133";
        internal const string MeshThemeHyperlink = "F28943";
        internal const string MeshThemeFollowedHyperlinkColor = "F1B76C";

        internal const string MetropolitanThemeName = "SpreadsheetLight Metropolitan";
        internal const string MetropolitanThemeMajorLatinFont = "Calibri Light";
        internal const string MetropolitanThemeMinorLatinFont = "Calibri Light";
        internal const string MetropolitanThemeDark1Color = "000000";
        internal const string MetropolitanThemeLight1Color = "FFFFFF";
        internal const string MetropolitanThemeDark2Color = "162F33";
        internal const string MetropolitanThemeLight2Color = "EAF0E0";
        internal const string MetropolitanThemeAccent1Color = "50B4C8";
        internal const string MetropolitanThemeAccent2Color = "A8B97F";
        internal const string MetropolitanThemeAccent3Color = "9B9256";
        internal const string MetropolitanThemeAccent4Color = "657689";
        internal const string MetropolitanThemeAccent5Color = "7A855D";
        internal const string MetropolitanThemeAccent6Color = "84AC9D";
        internal const string MetropolitanThemeHyperlink = "2370CD";
        internal const string MetropolitanThemeFollowedHyperlinkColor = "877589";

        internal const string MylarThemeName = "SpreadsheetLight Mylar";
        internal const string MylarThemeMajorLatinFont = "Corbel";
        internal const string MylarThemeMinorLatinFont = "Corbel";
        internal const string MylarThemeDark1Color = "000000";
        internal const string MylarThemeLight1Color = "FFFFFF";
        internal const string MylarThemeDark2Color = "656162";
        internal const string MylarThemeLight2Color = "E0DACC";
        internal const string MylarThemeAccent1Color = "4A5A7A";
        internal const string MylarThemeAccent2Color = "F7BD40";
        internal const string MylarThemeAccent3Color = "975C00";
        internal const string MylarThemeAccent4Color = "754D41";
        internal const string MylarThemeAccent5Color = "838995";
        internal const string MylarThemeAccent6Color = "687B66";
        internal const string MylarThemeHyperlink = "B5740B";
        internal const string MylarThemeFollowedHyperlinkColor = "7483A0";

        internal const string ParallaxThemeName = "SpreadsheetLight Parallax";
        internal const string ParallaxThemeMajorLatinFont = "Corbel";
        internal const string ParallaxThemeMinorLatinFont = "Corbel";
        internal const string ParallaxThemeDark1Color = "000000";
        internal const string ParallaxThemeLight1Color = "FFFFFF";
        internal const string ParallaxThemeDark2Color = "212121";
        internal const string ParallaxThemeLight2Color = "CDD0D1";
        internal const string ParallaxThemeAccent1Color = "30ACEC";
        internal const string ParallaxThemeAccent2Color = "80C34F";
        internal const string ParallaxThemeAccent3Color = "E29D3E";
        internal const string ParallaxThemeAccent4Color = "D64A3B";
        internal const string ParallaxThemeAccent5Color = "D64787";
        internal const string ParallaxThemeAccent6Color = "A666E1";
        internal const string ParallaxThemeHyperlink = "3085ED";
        internal const string ParallaxThemeFollowedHyperlinkColor = "82B6F4";

        internal const string QuotableThemeName = "SpreadsheetLight Quotable";
        internal const string QuotableThemeMajorLatinFont = "Century Gothic";
        internal const string QuotableThemeMinorLatinFont = "Century Gothic";
        internal const string QuotableThemeDark1Color = "000000";
        internal const string QuotableThemeLight1Color = "FFFFFF";
        internal const string QuotableThemeDark2Color = "212121";
        internal const string QuotableThemeLight2Color = "636363";
        internal const string QuotableThemeAccent1Color = "00C6BB";
        internal const string QuotableThemeAccent2Color = "6FEBA0";
        internal const string QuotableThemeAccent3Color = "B6DF5E";
        internal const string QuotableThemeAccent4Color = "EFB251";
        internal const string QuotableThemeAccent5Color = "EF755F";
        internal const string QuotableThemeAccent6Color = "ED515C";
        internal const string QuotableThemeHyperlink = "8F8F8F";
        internal const string QuotableThemeFollowedHyperlinkColor = "A5A5A5";

        internal const string SavonThemeName = "SpreadsheetLight Savon";
        internal const string SavonThemeMajorLatinFont = "Century Gothic";
        internal const string SavonThemeMinorLatinFont = "Century Gothic";
        internal const string SavonThemeDark1Color = "000000";
        internal const string SavonThemeLight1Color = "FFFFFF";
        internal const string SavonThemeDark2Color = "1485A4";
        internal const string SavonThemeLight2Color = "E3DED1";
        internal const string SavonThemeAccent1Color = "1CADE4";
        internal const string SavonThemeAccent2Color = "2683C6";
        internal const string SavonThemeAccent3Color = "27CED7";
        internal const string SavonThemeAccent4Color = "42BA97";
        internal const string SavonThemeAccent5Color = "3E8853";
        internal const string SavonThemeAccent6Color = "62A39F";
        internal const string SavonThemeHyperlink = "F49100";
        internal const string SavonThemeFollowedHyperlinkColor = "739D9B";

        internal const string SketchbookThemeName = "SpreadsheetLight Sketchbook";
        internal const string SketchbookThemeMajorLatinFont = "Cambria";
        internal const string SketchbookThemeMinorLatinFont = "Cambria";
        internal const string SketchbookThemeDark1Color = "000000";
        internal const string SketchbookThemeLight1Color = "FFFFFF";
        internal const string SketchbookThemeDark2Color = "4C1304";
        internal const string SketchbookThemeLight2Color = "FFFEE6";
        internal const string SketchbookThemeAccent1Color = "A63212";
        internal const string SketchbookThemeAccent2Color = "E68230";
        internal const string SketchbookThemeAccent3Color = "9BB05E";
        internal const string SketchbookThemeAccent4Color = "6B9BC7";
        internal const string SketchbookThemeAccent5Color = "4E66B2";
        internal const string SketchbookThemeAccent6Color = "8976AC";
        internal const string SketchbookThemeHyperlink = "942408";
        internal const string SketchbookThemeFollowedHyperlinkColor = "B34F17";

        internal const string SlateThemeName = "SpreadsheetLight Slate";
        internal const string SlateThemeMajorLatinFont = "Calisto MT";
        internal const string SlateThemeMinorLatinFont = "Calisto MT";
        internal const string SlateThemeDark1Color = "000000";
        internal const string SlateThemeLight1Color = "FFFFFF";
        internal const string SlateThemeDark2Color = "212123";
        internal const string SlateThemeLight2Color = "DADADA";
        internal const string SlateThemeAccent1Color = "BC451B";
        internal const string SlateThemeAccent2Color = "D3BA68";
        internal const string SlateThemeAccent3Color = "BB8640";
        internal const string SlateThemeAccent4Color = "AD9277";
        internal const string SlateThemeAccent5Color = "A55A43";
        internal const string SlateThemeAccent6Color = "AD9D7B";
        internal const string SlateThemeHyperlink = "E98052";
        internal const string SlateThemeFollowedHyperlinkColor = "F4B69B";

        internal const string SohoThemeName = "SpreadsheetLight Soho";
        internal const string SohoThemeMajorLatinFont = "Candara";
        internal const string SohoThemeMinorLatinFont = "Candara";
        internal const string SohoThemeDark1Color = "2E2224";
        internal const string SohoThemeLight1Color = "FFFFFF";
        internal const string SohoThemeDark2Color = "48231E";
        internal const string SohoThemeLight2Color = "CBD8DD";
        internal const string SohoThemeAccent1Color = "61625E";
        internal const string SohoThemeAccent2Color = "964D2C";
        internal const string SohoThemeAccent3Color = "66553E";
        internal const string SohoThemeAccent4Color = "848058";
        internal const string SohoThemeAccent5Color = "AFA14B";
        internal const string SohoThemeAccent6Color = "AD7D4D";
        internal const string SohoThemeHyperlink = "FFDE66";
        internal const string SohoThemeFollowedHyperlinkColor = "C0AEBC";

        internal const string SpringThemeName = "SpreadsheetLight Spring";
        internal const string SpringThemeMajorLatinFont = "Verdana";
        internal const string SpringThemeMinorLatinFont = "Verdana";
        internal const string SpringThemeDark1Color = "000000";
        internal const string SpringThemeLight1Color = "FFFFFF";
        internal const string SpringThemeDark2Color = "66822D";
        internal const string SpringThemeLight2Color = "BEEA73";
        internal const string SpringThemeAccent1Color = "C1EC76";
        internal const string SpringThemeAccent2Color = "8FE28A";
        internal const string SpringThemeAccent3Color = "F3BF45";
        internal const string SpringThemeAccent4Color = "F47E5A";
        internal const string SpringThemeAccent5Color = "F489CF";
        internal const string SpringThemeAccent6Color = "B56FF4";
        internal const string SpringThemeHyperlink = "408080";
        internal const string SpringThemeFollowedHyperlinkColor = "5EAEAE";

        internal const string SummerThemeName = "SpreadsheetLight Summer";
        internal const string SummerThemeMajorLatinFont = "Verdana";
        internal const string SummerThemeMinorLatinFont = "Verdana";
        internal const string SummerThemeDark1Color = "000000";
        internal const string SummerThemeLight1Color = "FFFFFF";
        internal const string SummerThemeDark2Color = "E89117";
        internal const string SummerThemeLight2Color = "FEDD78";
        internal const string SummerThemeAccent1Color = "A1B633";
        internal const string SummerThemeAccent2Color = "C4D73F";
        internal const string SummerThemeAccent3Color = "FFCE2D";
        internal const string SummerThemeAccent4Color = "FFA600";
        internal const string SummerThemeAccent5Color = "ED5E00";
        internal const string SummerThemeAccent6Color = "C62D03";
        internal const string SummerThemeHyperlink = "408080";
        internal const string SummerThemeFollowedHyperlinkColor = "5EAEAE";

        internal const string ThermalThemeName = "SpreadsheetLight Thermal";
        internal const string ThermalThemeMajorLatinFont = "Calibri";
        internal const string ThermalThemeMinorLatinFont = "Calibri";
        internal const string ThermalThemeDark1Color = "4D5B6B";
        internal const string ThermalThemeLight1Color = "FFFFFF";
        internal const string ThermalThemeDark2Color = "675D59";
        internal const string ThermalThemeLight2Color = "E8DED8";
        internal const string ThermalThemeAccent1Color = "FF7605";
        internal const string ThermalThemeAccent2Color = "7F7F7F";
        internal const string ThermalThemeAccent3Color = "7F5185";
        internal const string ThermalThemeAccent4Color = "89AAD3";
        internal const string ThermalThemeAccent5Color = "8F5B4B";
        internal const string ThermalThemeAccent6Color = "C84340";
        internal const string ThermalThemeHyperlink = "89AAD3";
        internal const string ThermalThemeFollowedHyperlinkColor = "795185";

        internal const string TradeshowThemeName = "SpreadsheetLight Tradeshow";
        internal const string TradeshowThemeMajorLatinFont = "Arial Black";
        internal const string TradeshowThemeMinorLatinFont = "Candara";
        internal const string TradeshowThemeDark1Color = "3F3F3F";
        internal const string TradeshowThemeLight1Color = "FFFFFF";
        internal const string TradeshowThemeDark2Color = "7DAFC3";
        internal const string TradeshowThemeLight2Color = "E5E4DF";
        internal const string TradeshowThemeAccent1Color = "7C959A";
        internal const string TradeshowThemeAccent2Color = "DB8631";
        internal const string TradeshowThemeAccent3Color = "E3CC5A";
        internal const string TradeshowThemeAccent4Color = "ACADA8";
        internal const string TradeshowThemeAccent5Color = "927C61";
        internal const string TradeshowThemeAccent6Color = "B3B435";
        internal const string TradeshowThemeHyperlink = "0079A4";
        internal const string TradeshowThemeFollowedHyperlinkColor = "595959";

        internal const string UrbanPopThemeName = "SpreadsheetLight UrbanPop";
        internal const string UrbanPopThemeMajorLatinFont = "Gill Sans MT";
        internal const string UrbanPopThemeMinorLatinFont = "Gill Sans MT";
        internal const string UrbanPopThemeDark1Color = "000000";
        internal const string UrbanPopThemeLight1Color = "FFFFFF";
        internal const string UrbanPopThemeDark2Color = "282828";
        internal const string UrbanPopThemeLight2Color = "D4D4D4";
        internal const string UrbanPopThemeAccent1Color = "86CE24";
        internal const string UrbanPopThemeAccent2Color = "00A2E6";
        internal const string UrbanPopThemeAccent3Color = "FAC810";
        internal const string UrbanPopThemeAccent4Color = "7D8F8C";
        internal const string UrbanPopThemeAccent5Color = "D06B20";
        internal const string UrbanPopThemeAccent6Color = "958B8B";
        internal const string UrbanPopThemeHyperlink = "FF9900";
        internal const string UrbanPopThemeFollowedHyperlinkColor = "969696";

        internal const string VaporTrailThemeName = "SpreadsheetLight Vapor Trail";
        internal const string VaporTrailThemeMajorLatinFont = "Century Gothic";
        internal const string VaporTrailThemeMinorLatinFont = "Century Gothic";
        internal const string VaporTrailThemeDark1Color = "000000";
        internal const string VaporTrailThemeLight1Color = "FFFFFF";
        internal const string VaporTrailThemeDark2Color = "454545";
        internal const string VaporTrailThemeLight2Color = "DADADA";
        internal const string VaporTrailThemeAccent1Color = "DF2E28";
        internal const string VaporTrailThemeAccent2Color = "FE801A";
        internal const string VaporTrailThemeAccent3Color = "E9BF35";
        internal const string VaporTrailThemeAccent4Color = "81BB42";
        internal const string VaporTrailThemeAccent5Color = "32C7A9";
        internal const string VaporTrailThemeAccent6Color = "4A9BDC";
        internal const string VaporTrailThemeHyperlink = "F0532B";
        internal const string VaporTrailThemeFollowedHyperlinkColor = "F38B53";

        internal const string ViewThemeName = "SpreadsheetLight View";
        internal const string ViewThemeMajorLatinFont = "Century Schoolbook";
        internal const string ViewThemeMinorLatinFont = "Century Schoolbook";
        internal const string ViewThemeDark1Color = "000000";
        internal const string ViewThemeLight1Color = "FFFFFF";
        internal const string ViewThemeDark2Color = "46464A";
        internal const string ViewThemeLight2Color = "D6D3CC";
        internal const string ViewThemeAccent1Color = "6F6F74";
        internal const string ViewThemeAccent2Color = "92A9B9";
        internal const string ViewThemeAccent3Color = "A7B789";
        internal const string ViewThemeAccent4Color = "B9A489";
        internal const string ViewThemeAccent5Color = "8D6374";
        internal const string ViewThemeAccent6Color = "9B7362";
        internal const string ViewThemeHyperlink = "67AABF";
        internal const string ViewThemeFollowedHyperlinkColor = "ABAFA5";

        internal const string WinterThemeName = "SpreadsheetLight Winter";
        internal const string WinterThemeMajorLatinFont = "Verdana";
        internal const string WinterThemeMinorLatinFont = "Verdana";
        internal const string WinterThemeDark1Color = "000000";
        internal const string WinterThemeLight1Color = "FFFFFF";
        internal const string WinterThemeDark2Color = "1F7BB6";
        internal const string WinterThemeLight2Color = "C5E1FE";
        internal const string WinterThemeAccent1Color = "B2BDC1";
        internal const string WinterThemeAccent2Color = "767D83";
        internal const string WinterThemeAccent3Color = "3E505C";
        internal const string WinterThemeAccent4Color = "386489";
        internal const string WinterThemeAccent5Color = "4C80AF";
        internal const string WinterThemeAccent6Color = "7DA7D1";
        internal const string WinterThemeHyperlink = "408080";
        internal const string WinterThemeFollowedHyperlinkColor = "5EAEAE";

        internal const string WoodTypeThemeName = "SpreadsheetLight Wood Type";
        internal const string WoodTypeThemeMajorLatinFont = "Rockwell Condensed";
        internal const string WoodTypeThemeMinorLatinFont = "Rockwell";
        internal const string WoodTypeThemeDark1Color = "000000";
        internal const string WoodTypeThemeLight1Color = "FFFFFF";
        internal const string WoodTypeThemeDark2Color = "696464";
        internal const string WoodTypeThemeLight2Color = "E9E5DC";
        internal const string WoodTypeThemeAccent1Color = "D34817";
        internal const string WoodTypeThemeAccent2Color = "9B2D1F";
        internal const string WoodTypeThemeAccent3Color = "A28E6A";
        internal const string WoodTypeThemeAccent4Color = "956251";
        internal const string WoodTypeThemeAccent5Color = "918485";
        internal const string WoodTypeThemeAccent6Color = "855D5D";
        internal const string WoodTypeThemeHyperlink = "CC9900";
        internal const string WoodTypeThemeFollowedHyperlinkColor = "96A9A9";

        // not exactly constants? I don't care. They look like constants when you're typing.
        // That's good enough for me...
        internal static DateTime Epoch1900()
        {
            return new DateTime(1900, 1, 1, 0, 0, 0, 0);
        }

        internal static DateTime Epoch1904()
        {
             return new DateTime(1904, 1, 1, 0, 0, 0, 0);
        }
    }
}
