/*
Copyright (c) 2011 Vincent Tan Wai Lip

Permission is hereby granted, free of charge, to any person obtaining a copy of this software
and associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

*/

// You can get the precompiled DLL as well as documentation at http://spreadsheetlight.com/
// If you're interested, my personal blog is at http://polymathprogrammer.com/
// Thanks for using SpreadsheetLight! -Vincent

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SpreadsheetLight;

/// <summary>
/// Encapsulates the main properties and methods to create and manipulate a spreadsheet.
/// </summary>
public partial class SLDocument : IDisposable
{
    private MemoryStream memstream;
    private SpreadsheetDocument xl;
    internal WorkbookPart wbp;
    private string gsSpreadsheetFileName = string.Empty;
    private uint giSelectedWorksheetID = 0;
    internal string gsSelectedWorksheetName = string.Empty;
    private string gsSelectedWorksheetRelationshipID = string.Empty;
    private int giWorksheetIdCounter = 1;
    private bool IsNewSpreadsheet = true;
    private bool IsNewWorksheet = true;

    private bool gbThrowExceptionsIfAny = false;
    // This is a suggestion from David H. SpreadsheetLight originally swallows any exceptions
    // so that the calling program using the library continues to function.
    // It's designed this way because the calling program can be an automated scheduled program
    // running in the wee hours of the morning. If there's a whole bunch of programs, then a failure
    // can mean the whole schedule is screwed up. I'm not ready for that kind of responsibility...
    // However, David suggested that maybe make it such that the developer (that's you) is allowed
    // to take on this responsibility. Thus I'm exposing this property. You've been warned.
    /// <summary>
    /// Set to true to bubble exceptions up if there are any occurring within SpreadsheetLight. Set to false otherwise. The default is false.
    /// </summary>
    public bool ThrowExceptionsIfAny
    {
        get { return gbThrowExceptionsIfAny; }
        set { gbThrowExceptionsIfAny = value; }
    }

    /// <summary>
    /// The file name of the default blank spreadsheet. This is read-only.
    /// </summary>
    public static string BlankSpreadsheetFileName
    {
        get { return SLConstants.BlankSpreadsheetFileName; }
    }

    /// <summary>
    /// The name of the first worksheet when creating a new spreadsheet. This is read-only.
    /// </summary>
    public static string DefaultFirstSheetName
    {
        get { return SLConstants.DefaultFirstSheetName; }
    }

    /// <summary>
    /// The default font size in points. This is read-only.
    /// </summary>
    public static double DefaultFontSize
    {
        get { return SLConstants.DefaultFontSize; }
    }

    /// <summary>
    /// The maximum number of rows in a worksheet. This is read-only.
    /// </summary>
    public static int RowLimit
    {
        get { return SLConstants.RowLimit; }
    }

    /// <summary>
    /// The maximum number of columns in a worksheet. This is read-only.
    /// </summary>
    public static int ColumnLimit
    {
        get { return SLConstants.ColumnLimit; }
    }

    // this is assuming a default 96 DPI
    private static long lPixelToEMU = SLConstants.InchToEMU / 96;
    /// <summary>
    /// This should be assigned immediately after the minor font is determined
    /// and never assigned anymore after that.
    /// </summary>
    internal static long PixelToEMU
    {
        get { return lPixelToEMU; }
        set { lPixelToEMU = value; }
    }

    // this is assuming a default 96 DPI
    // meaning 1 pixel height multiple = 1 pixel * 72 (points per inch) / 96 (DPI)
    private static double fRowHeightMultiple = 0.75;
    /// <summary>
    /// This should be assigned immediately after the minor font is determined
    /// and never assigned anymore after that.
    /// </summary>
    internal static double RowHeightMultiple
    {
        get { return fRowHeightMultiple; }
        set { fRowHeightMultiple = value; }
    }

    /// <summary>
    /// The document metadata.
    /// </summary>
    public SLDocumentProperties DocumentProperties { get; set; }

    // key: shared string index, separator, style index
    internal Dictionary<string, SizeF> dictAutoFitSharedStringCache;

    // key: plain text string, separator, style index
    // Ideally, it should be the font index and not style index for better efficiency.
    // But at the point of autofitting, the font index is not available.
    // And it doesn't seem worth it to calculate the font index.
    internal Dictionary<string, SizeF> dictAutoFitTextCache;

    internal Dictionary<int, System.Drawing.Font> dictAutoFitFontCache;
    internal Dictionary<int, string> dictAutoFitFormatCodeCache;
    internal Dictionary<int, int> dictAutoFitTextRotationCache;

    internal SLWorkbook slwb;
    internal SLWorksheet slws;

    internal SLSimpleTheme SimpleTheme;

    internal int NumberFormatGeneralId;
    internal string NumberFormatGeneralText;
    internal int NextNumberFormatId;
    internal Dictionary<int, string> dictBuiltInNumberingFormat;
    internal Dictionary<string, int> dictBuiltInNumberingFormatHash;

    internal Dictionary<int, string> dictStyleNumberingFormat;
    internal Dictionary<string, int> dictStyleNumberingFormatHash;

    internal int countStyle;
    internal List<string> listStyle;
    internal Dictionary<string, int> dictStyleHash;

    internal int countStyleFont;
    internal List<string> listStyleFont;
    internal Dictionary<string, int> dictStyleFontHash;

    internal int countStyleFill;
    internal List<string> listStyleFill;
    internal Dictionary<string, int> dictStyleFillHash;

    internal int countStyleBorder;
    internal List<string> listStyleBorder;
    internal Dictionary<string, int> dictStyleBorderHash;

    internal int countStyleCellStyleFormat;
    internal List<string> listStyleCellStyleFormat;
    internal Dictionary<string, int> dictStyleCellStyleFormatHash;

    internal int countStyleCellStyle;
    internal List<string> listStyleCellStyle;
    internal Dictionary<string, int> dictStyleCellStyleHash;

    internal int countStyleDifferentialFormat;
    internal List<string> listStyleDifferentialFormat;
    internal Dictionary<string, int> dictStyleDifferentialFormatHash;

    internal int countStyleTableStyle;
    internal List<string> listStyleTableStyle;
    internal Dictionary<string, int> dictStyleTableStyleHash;

    internal Colors StylesheetColors = null;

    private string TableStylesDefaultTableStyle = string.Empty;
    private string TableStylesDefaultPivotStyle = string.Empty;

    internal int countSharedString;
    internal List<string> listSharedString;
    internal Dictionary<string, int> dictSharedStringHash;

    internal HashSet<string> hsUniqueSharedString;
    private bool gbWriteUniqueSharedStringCount = true;

    /// <summary>
    /// Set to true to write the unique shared string count property, false otherwise. The default is true.
    /// Set this to false if you get an error opening the resulting file because at a high number of text strings,
    /// SpreadsheetLight might count the number wrongly (because the Open XML specs are sort of convoluted).
    /// </summary>
    public bool WriteUniqueSharedStringCount
    {
        get { return gbWriteUniqueSharedStringCount; }
        set { gbWriteUniqueSharedStringCount = value; }
    }


    /// <summary>
    /// Create a new spreadsheet with a worksheet with the default sheet name.
    /// </summary>
    public SLDocument()
    {
        memstream = new MemoryStream();
        xl = SpreadsheetDocument.Create(memstream, SpreadsheetDocumentType.Workbook);
        wbp = xl.AddWorkbookPart();
        IsNewSpreadsheet = true;
        slwb = new SLWorkbook();

        this.DocumentProperties = new SLDocumentProperties();
        this.DocumentProperties.Created = DateTime.UtcNow.ToString(SLConstants.W3CDTF);

        InitialiseAutoFitCache();

        LoadBuiltInNumberingFormats();
        InitialiseStylesheetWhatNots(SLThemeTypeValues.Office);
        LoadSharedStringTable();

        InitialiseNewSpreadsheet();
    }

    /// <summary>
    /// Create a new spreadsheet with a selected theme of fonts and colors.
    /// </summary>
    /// <param name="ThemeType">The selected theme.</param>
    public SLDocument(SLThemeTypeValues ThemeType)
    {
        memstream = new MemoryStream();
        xl = SpreadsheetDocument.Create(memstream, SpreadsheetDocumentType.Workbook);
        wbp = xl.AddWorkbookPart();
        IsNewSpreadsheet = true;
        slwb = new SLWorkbook();

        this.DocumentProperties = new SLDocumentProperties();
        this.DocumentProperties.Created = DateTime.UtcNow.ToString(SLConstants.W3CDTF);

        InitialiseAutoFitCache();

        LoadBuiltInNumberingFormats();
        InitialiseStylesheetWhatNots(ThemeType);
        LoadSharedStringTable();

        InitialiseNewSpreadsheet();
    }

    /// <summary>
    /// Create a new spreadsheet with a custom theme.
    /// </summary>
    /// <param name="ThemeSettings">Custom theme settings.</param>
    public SLDocument(SLThemeSettings ThemeSettings)
    {
        memstream = new MemoryStream();
        xl = SpreadsheetDocument.Create(memstream, SpreadsheetDocumentType.Workbook);
        wbp = xl.AddWorkbookPart();
        IsNewSpreadsheet = true;
        slwb = new SLWorkbook();

        this.DocumentProperties = new SLDocumentProperties();
        this.DocumentProperties.Created = DateTime.UtcNow.ToString(SLConstants.W3CDTF);

        InitialiseAutoFitCache();

        LoadBuiltInNumberingFormats();
        InitialiseStylesheetWhatNots(ThemeSettings);
        LoadSharedStringTable();

        InitialiseNewSpreadsheet();
    }

    private void InitialiseNewSpreadsheet()
    {
        gsSpreadsheetFileName = SLConstants.BlankSpreadsheetFileName;
        gsSelectedWorksheetName = SLConstants.DefaultFirstSheetName;
        giWorksheetIdCounter = 0;

        // the theme should be loaded by now.
        // That means the default row height and column widths are also calculated.
        AddWorksheet(gsSelectedWorksheetName);
    }

    /// <summary>
    /// Open an existing spreadsheet, with the first available worksheet loaded.
    /// Note that the first available worksheet may not be visible, nor in the order you think it is.
    /// </summary>
    /// <param name="SpreadsheetFileName">The file name of the existing spreadsheet.</param>
    public SLDocument(string SpreadsheetFileName)
    {
        byte[] baData = File.ReadAllBytes(SpreadsheetFileName);
        memstream = new MemoryStream();
        memstream.Write(baData, 0, baData.Length);

        gsSpreadsheetFileName = SpreadsheetFileName;

        OpenExistingSpreadsheet(string.Empty);
    }

    /// <summary>
    /// Open an existing spreadsheet, with the desired worksheet ready for use.
    /// This optimizes loading so the desired worksheet's contents are loaded directly instead of first loading the first available worksheet.
    /// Note that if the given sheet name doesn't exist, the first available worksheet is loaded.
    /// </summary>
    /// <param name="SpreadsheetFileName">The file name of the existing spreadsheet.</param>
    /// <param name="SheetNameOnOpen">The sheet name of desired worksheet on opening the spreadsheet.</param>
    public SLDocument(string SpreadsheetFileName, string SheetNameOnOpen)
    {
        byte[] baData = File.ReadAllBytes(SpreadsheetFileName);
        memstream = new MemoryStream();
        memstream.Write(baData, 0, baData.Length);

        gsSpreadsheetFileName = SpreadsheetFileName;

        OpenExistingSpreadsheet(SheetNameOnOpen);
    }

    /// <summary>
    /// Open an existing spreadsheet from a Stream, with the first available worksheet loaded.
    /// Note that the first available worksheet may not be visible, nor in the order you think it is.
    /// </summary>
    /// <param name="SpreadsheetStream">Stream containing spreadsheet content.</param>
    public SLDocument(Stream SpreadsheetStream)
    {
        SpreadsheetStream.Position = 0;
        byte[] baData = new byte[SpreadsheetStream.Length];
        SpreadsheetStream.Read(baData, 0, baData.Length);
        memstream = new MemoryStream();
        memstream.Write(baData, 0, baData.Length);

        gsSpreadsheetFileName = SLConstants.BlankSpreadsheetFileName;

        OpenExistingSpreadsheet(string.Empty);
    }

    /// <summary>
    /// Open an existing spreadsheet from a Stream, with the desired worksheet ready for use.
    /// This optimizes loading so the desired worksheet's contents are loaded directly instead of first loading the first available worksheet.
    /// Note that if the given sheet name doesn't exist, the first available worksheet is loaded.
    /// </summary>
    /// <param name="SpreadsheetStream">Stream containing spreadsheet content.</param>
    /// <param name="SheetNameOnOpen">The sheet name of desired worksheet on opening the spreadsheet.</param>
    public SLDocument(Stream SpreadsheetStream, string SheetNameOnOpen)
    {
        SpreadsheetStream.Position = 0;
        byte[] baData = new byte[SpreadsheetStream.Length];
        SpreadsheetStream.Read(baData, 0, baData.Length);
        memstream = new MemoryStream();
        memstream.Write(baData, 0, baData.Length);

        gsSpreadsheetFileName = SLConstants.BlankSpreadsheetFileName;

        OpenExistingSpreadsheet(SheetNameOnOpen);
    }

    private void OpenExistingSpreadsheet(string SheetNameOnOpen)
    {
        xl = SpreadsheetDocument.Open(memstream, true);
        wbp = xl.WorkbookPart;
        IsNewSpreadsheet = false;
        slwb = new SLWorkbook();

        this.DocumentProperties = new SLDocumentProperties();
        this.LoadDocumentProperties();

        InitialiseAutoFitCache();

        LoadBuiltInNumberingFormats();
        InitialiseStylesheetWhatNots(SLThemeTypeValues.Office);
        LoadSharedStringTable();

        giWorksheetIdCounter = 0;
        using (OpenXmlReader oxr = OpenXmlReader.Create(wbp))
        {
            SLWorkbookView wv;
            Sheet s;
            SLSheet sheet;
            DefinedName dn;
            SLDefinedName sldn;
            while (oxr.Read())
            {
                if (oxr.ElementType == typeof(WorkbookView))
                {
                    wv = new SLWorkbookView();
                    wv.FromWorkbookView((WorkbookView)oxr.LoadCurrentElement());
                    slwb.WorkbookViews.Add(wv);
                }
                else if (oxr.ElementType == typeof(Sheet))
                {
                    s = (Sheet)oxr.LoadCurrentElement();
                    sheet = new SLSheet(s.Name.Value, s.SheetId.Value, s.Id.Value, SLSheetType.Unknown);
                    if (s.State != null) sheet.State = s.State.Value;
                    slwb.Sheets.Add(sheet);
                    if (sheet.SheetId > giWorksheetIdCounter)
                    {
                        giWorksheetIdCounter = (int)sheet.SheetId;
                    }
                }
                else if (oxr.ElementType == typeof(DefinedName))
                {
                    dn = (DefinedName)oxr.LoadCurrentElement();
                    sldn = new SLDefinedName(dn.Name.Value);
                    sldn.FromDefinedName(dn);
                    slwb.DefinedNames.Add(sldn);
                }
                else if (oxr.ElementType == typeof(PivotCache))
                {
                    // cache IDs supposed to be unique, so I'm not gonna check for the hash set
                    slwb.PivotTableCacheIds.Add(((PivotCache)oxr.LoadCurrentElement()).CacheId.Value);
                }
            }
        }

        if (wbp.Workbook.WorkbookProperties != null)
        {
            slwb.WorkbookProperties.FromWorkbookProperties(wbp.Workbook.WorkbookProperties);
        }

        if (wbp.CalculationChainPart != null)
        {
            int iCurrentSheetId = 0;
            SLCalculationCell slcc = new SLCalculationCell(string.Empty);
            CalculationCell cc;
            using (OpenXmlReader oxr = OpenXmlReader.Create(wbp.CalculationChainPart))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(CalculationCell))
                    {
                        cc = (CalculationCell)oxr.LoadCurrentElement();
                        if (cc.SheetId == null)
                        {
                            cc.SheetId = iCurrentSheetId;
                        }
                        else
                        {
                            if (cc.SheetId.Value != iCurrentSheetId)
                                iCurrentSheetId = cc.SheetId.Value;
                        }
                        slcc.FromCalculationCell(cc);
                        slwb.CalculationCells.Add(slcc.Clone());
                    }
                }
            }
        }

        // To determine the type of sheet. Do this before the part
        // where the table and pivot table parts are set.

        bool bFound = false;
        string sRelID = string.Empty;
        foreach (SLSheet sheet in slwb.Sheets)
        {
            bFound = false;
            foreach (WorksheetPart wspFound in wbp.WorksheetParts)
            {
                sRelID = wbp.GetIdOfPart(wspFound);
                if (sheet.Id.Equals(sRelID, StringComparison.OrdinalIgnoreCase))
                {
                    sheet.SheetType = SLSheetType.Worksheet;
                    bFound = true;
                    break;
                }
            }

            if (!bFound)
            {
                foreach (ChartsheetPart csp in wbp.ChartsheetParts)
                {
                    sRelID = wbp.GetIdOfPart(csp);
                    if (sheet.Id.Equals(sRelID, StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.SheetType = SLSheetType.Chartsheet;
                        bFound = true;
                        break;
                    }
                }
            }

            if (!bFound)
            {
                foreach (DialogsheetPart dsp in wbp.DialogsheetParts)
                {
                    sRelID = wbp.GetIdOfPart(dsp);
                    if (sheet.Id.Equals(sRelID, StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.SheetType = SLSheetType.DialogSheet;
                        bFound = true;
                        break;
                    }
                }
            }

            if (!bFound)
            {
                foreach (MacroSheetPart msp in wbp.MacroSheetParts)
                {
                    sRelID = wbp.GetIdOfPart(msp);
                    if (sheet.Id.Equals(sRelID, StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.SheetType = SLSheetType.Macrosheet;
                        bFound = true;
                        break;
                    }
                }
            }
        }

        WorksheetPart wsp;
        foreach (SLSheet sheet in slwb.Sheets)
        {
            if (sheet.SheetType == SLSheetType.Worksheet)
            {
                wsp = (WorksheetPart)wbp.GetPartById(sheet.Id);
                foreach (TableDefinitionPart tdp in wsp.TableDefinitionParts)
                {
                    if (tdp.Table.Id != null && !slwb.TableIds.Contains(tdp.Table.Id.Value))
                        slwb.TableIds.Add(tdp.Table.Id.Value);

                    if (tdp.Table.Name != null && !slwb.TableNames.Contains(tdp.Table.Name.Value))
                        slwb.TableNames.Add(tdp.Table.Name.Value);
                }

                foreach (PivotTablePart ptp in wsp.PivotTableParts)
                {
                    if (ptp.PivotTableDefinition.Name != null
                        && !slwb.PivotTableNames.Contains(ptp.PivotTableDefinition.Name.Value))
                        slwb.PivotTableNames.Add(ptp.PivotTableDefinition.Name.Value);

                    // the cache ID should already be added, from the workbook part above.
                    // But we check again just to be sure. Cache IDs have to be unique throughout
                    // the workbook.
                    if (ptp.PivotTableDefinition.CacheId != null
                        && !slwb.PivotTableCacheIds.Contains(ptp.PivotTableDefinition.CacheId.Value))
                        slwb.PivotTableCacheIds.Add(ptp.PivotTableDefinition.CacheId.Value);
                }
            }
        }

        string sWorksheetName = SLConstants.DefaultFirstSheetName;
        int i = 1;
        bool bCannotFind = true;
        bool bIsLegit = true;
        if (wbp.WorksheetParts.Count() == 0)
        {
            // no worksheets! Apparently an Excel file with only 1 dialog sheet is perfectly legit...
            // come up with a legit worksheet name that's not already taken...
            i = 1;
            bCannotFind = true;
            while (bCannotFind)
            {
                sWorksheetName = string.Format("Sheet{0}", i);
                bIsLegit = true;
                foreach (SLSheet sheet in slwb.Sheets)
                {
                    if (sheet.Name.Equals(sWorksheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        bIsLegit = false;
                        break;
                    }
                }
                ++i;
                if (bIsLegit) bCannotFind = false;
            }

            AddWorksheet(sWorksheetName);
        }
        else
        {
            bFound = false;
            // there's a given worksheet name
            if (SheetNameOnOpen.Length > 0)
            {
                foreach (SLSheet sheet in slwb.Sheets)
                {
                    if (sheet.Name.Equals(SheetNameOnOpen, StringComparison.OrdinalIgnoreCase)
                        && sheet.SheetType == SLSheetType.Worksheet)
                    {
                        giSelectedWorksheetID = sheet.SheetId;
                        gsSelectedWorksheetName = sheet.Name;
                        gsSelectedWorksheetRelationshipID = sheet.Id;
                        bFound = true;
                        break;
                    }
                }
            }
            else
            {
                // we try to get the "actively selected" worksheet already selected.
                uint iActiveTab = 0;
                if (slwb.WorkbookViews.Count > 0)
                {
                    iActiveTab = slwb.WorkbookViews[0].ActiveTab;
                }

                // there should be at least *this* number of sheets (whether it's a worksheet
                // chartsheet or whatnot).
                if (slwb.Sheets.Count > iActiveTab
                    && slwb.Sheets[(int)iActiveTab].SheetType == SLSheetType.Worksheet)
                {
                    giSelectedWorksheetID = slwb.Sheets[(int)iActiveTab].SheetId;
                    gsSelectedWorksheetName = slwb.Sheets[(int)iActiveTab].Name;
                    gsSelectedWorksheetRelationshipID = slwb.Sheets[(int)iActiveTab].Id;
                    bFound = true;
                }
            }

            if (!bFound)
            {
                // we get here either if there's no given worksheet name (bFound is still false),
                // or there's a given worksheet name but corresponding values weren't found.
                // The given worksheet name must be that of a worksheet. A chartsheet name is
                // considered "invalid".
                // Either way, we use the first available worksheet as the selected worksheet.
                wsp = wbp.WorksheetParts.First();
                sRelID = wbp.GetIdOfPart(wsp);

                foreach (SLSheet sheet in slwb.Sheets)
                {
                    if (sheet.Id.Equals(sRelID, StringComparison.OrdinalIgnoreCase))
                    {
                        giSelectedWorksheetID = sheet.SheetId;
                        gsSelectedWorksheetName = sheet.Name;
                        gsSelectedWorksheetRelationshipID = sheet.Id;
                        bFound = true;
                        break;
                    }
                }
            }

            if (bFound)
            {
                // A viable worksheet should be found by now. Otherwise, it's probably
                // a corrupted spreadsheet...
                LoadSelectedWorksheet();
                IsNewWorksheet = false;
            }
            else
            {
                // why is it not found!?! The file is corrupted somehow... we'll try to recover
                // by adding a new worksheet and selecting it. Same algorithm as above.
                i = 1;
                bCannotFind = true;
                while (bCannotFind)
                {
                    sWorksheetName = string.Format("Sheet{0}", i);
                    bIsLegit = true;
                    foreach (SLSheet sheet in slwb.Sheets)
                    {
                        if (sheet.Name.Equals(sWorksheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            bIsLegit = false;
                            break;
                        }
                    }
                    ++i;
                    if (bIsLegit) bCannotFind = false;
                }

                AddWorksheet(sWorksheetName);
            }
        }
    }

    private void InitialiseAutoFitCache()
    {
        this.dictAutoFitSharedStringCache = new Dictionary<string, SizeF>();
        this.dictAutoFitTextCache = new Dictionary<string, SizeF>();
        this.dictAutoFitFormatCodeCache = new Dictionary<int, string>();
        this.dictAutoFitTextRotationCache = new Dictionary<int, int>();

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            this.dictAutoFitFontCache = new Dictionary<int, System.Drawing.Font>();
    }

    private void InitialiseStylesheetWhatNots(SLThemeTypeValues ThemeType)
    {
        SimpleTheme = new SLSimpleTheme(wbp, ThemeType, gbThrowExceptionsIfAny);
        this.LoadStylesheet();
    }

    private void InitialiseStylesheetWhatNots(SLThemeSettings ThemeSettings)
    {
        SimpleTheme = new SLSimpleTheme(wbp, ThemeSettings, gbThrowExceptionsIfAny);
        this.LoadStylesheet();
    }

    private void LoadBuiltInNumberingFormats()
    {
        //http://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo.shortdatepattern.aspx
        DateTimeFormatInfo dtfi = CultureInfo.CurrentCulture.DateTimeFormat;

        dictBuiltInNumberingFormat = new Dictionary<int, string>();
        dictBuiltInNumberingFormat[0] = SLConstants.NumberFormatGeneral;
        dictBuiltInNumberingFormat[1] = "0";
        dictBuiltInNumberingFormat[2] = "0.00";
        dictBuiltInNumberingFormat[3] = "#,##0";
        dictBuiltInNumberingFormat[4] = "#,##0.00";
        dictBuiltInNumberingFormat[9] = "0%";
        dictBuiltInNumberingFormat[10] = "0.00%";
        dictBuiltInNumberingFormat[11] = "0.00E+00";
        dictBuiltInNumberingFormat[12] = "# ?/?";
        dictBuiltInNumberingFormat[13] = "# ??/??";

        // This is the format in the Open XML specs. But the specs say it can change:
        // "Some of these Ids can be interpreted differently, depending on the UI
        // language of the implementing application."
        //dictBuiltInNumberingFormat[14] = "mm-dd-yy";
        // So we try this:
        dictBuiltInNumberingFormat[14] = dtfi.ShortDatePattern.ToLowerInvariant();
        // We do ToLowerInvariant() because it returns something like "M/d/yyyy".
        // We want "m/d/yyyy" because the lower-case version is how Excel uses it.
        // Apparently, built-in index 14 is the short date version in Excel. Interesting...

        dictBuiltInNumberingFormat[15] = "d-mmm-yy";
        dictBuiltInNumberingFormat[16] = "d-mmm";
        dictBuiltInNumberingFormat[17] = "mmm-yy";
        dictBuiltInNumberingFormat[18] = "h:mm AM/PM";
        dictBuiltInNumberingFormat[19] = "h:mm:ss AM/PM";
        dictBuiltInNumberingFormat[20] = "h:mm";
        dictBuiltInNumberingFormat[21] = "h:mm:ss";
        dictBuiltInNumberingFormat[22] = "m/d/yy h:mm";
        dictBuiltInNumberingFormat[37] = "#,##0 ;(#,##0)";
        dictBuiltInNumberingFormat[38] = "#,##0 ;[Red](#,##0)";
        dictBuiltInNumberingFormat[39] = "#,##0.00;(#,##0.00)";
        dictBuiltInNumberingFormat[40] = "#,##0.00;[Red](#,##0.00)";
        dictBuiltInNumberingFormat[45] = "mm:ss";
        dictBuiltInNumberingFormat[46] = "[h]:mm:ss";
        dictBuiltInNumberingFormat[47] = "mmss.0";
        dictBuiltInNumberingFormat[48] = "##0.0E+0";
        dictBuiltInNumberingFormat[49] = "@";

        dictBuiltInNumberingFormatHash = new Dictionary<string, int>();
        foreach (var key in dictBuiltInNumberingFormat.Keys)
        {
            dictBuiltInNumberingFormatHash[dictBuiltInNumberingFormat[key]] = key;
        }
    }

    private void CreateFirstEmptyWorksheet()
    {
        gsSelectedWorksheetName = SLConstants.DefaultFirstSheetName;
        giWorksheetIdCounter = 0;

        AddWorksheet(gsSelectedWorksheetName);
    }

    private void LoadSelectedWorksheet()
    {
        // Need to check?
        //if (string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID)) return;

        WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);

        slws = new SLWorksheet(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors, SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize, SimpleTheme.ThemeRowHeight);
        
        int index = 0;
        SLColumnProperties cp;
        Column col;
        SLSheetView slsv;
        MergeCell mc;
        SLMergeCell slmc;
        SLConditionalFormatting condformat;
        SLHyperlink hl;

        OpenXmlReader oxrRow;
        int iRowIndex = -1;
        int iColumnIndex = -1;
        int iGuessRowIndex = 0;
        int iGuessColumnIndex = 0;
        SLRowProperties rp;
        Row r;
        Cell c;
        SLCell slc;
        SLSharedCellFormula scf;

        OpenXmlReader oxr = OpenXmlReader.Create(wsp);
        while (oxr.Read())
        {
            if (oxr.ElementType == typeof(SheetProperties))
            {
                SheetProperties sprop = (SheetProperties)oxr.LoadCurrentElement();
                slws.PageSettings.SheetProperties.FromSheetProperties(sprop);
            }
            else if (oxr.ElementType == typeof(SheetView))
            {
                slsv = new SLSheetView();
                slsv.FromSheetView((SheetView)oxr.LoadCurrentElement());
                slws.SheetViews.Add(slsv);

                // we're concerned only with the first workbook view.
                if (slsv.ShowFormulas && slsv.WorkbookViewId == 0) slws.IsDoubleColumnWidth = true;
            }
            else if (oxr.ElementType == typeof(SheetFormatProperties))
            {
                SheetFormatProperties sfp = (SheetFormatProperties)oxr.LoadCurrentElement();
                slws.SheetFormatProperties.FromSheetFormatProperties(sfp);
            }
            else if (oxr.ElementType == typeof(Column))
            {
                #region Column
                int i = 0;
                col = (Column)oxr.LoadCurrentElement();
                int min = (int)col.Min.Value;
                int max = (int)col.Max.Value;
                for (i = min; i <= max; ++i)
                {
                    cp = new SLColumnProperties(SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize);
                    if (col.Width != null)
                    {
                        cp.Width = col.Width.Value;
                        cp.HasWidth = true;
                    }
                    if (col.Style != null)
                    {
                        index = (int)col.Style.Value;
                        // default is 0
                        if (index > 0 && index < listStyle.Count)
                        {
                            cp.StyleIndex = (uint)index;
                        }
                    }
                    if (col.Hidden != null && col.Hidden.Value) cp.Hidden = col.Hidden.Value;
                    if (col.BestFit != null && col.BestFit.Value) cp.BestFit = col.BestFit.Value;
                    if (col.Phonetic != null && col.Phonetic.Value) cp.Phonetic = col.Phonetic.Value;
                    if (col.OutlineLevel != null && col.OutlineLevel.Value > 0) cp.OutlineLevel = col.OutlineLevel.Value;
                    if (col.Collapsed != null && col.Collapsed.Value) cp.Collapsed = col.Collapsed.Value;
                    slws.ColumnProperties[i] = cp;
                }
                #endregion
            }
            else if (oxr.ElementType == typeof(Row))
            {
                #region Row
                ++iGuessRowIndex;
                iGuessColumnIndex = 0;
                r = (Row)oxr.LoadCurrentElement();
                rp = new SLRowProperties(SimpleTheme.ThemeRowHeight);
                if (r.RowIndex != null)
                {
                    iRowIndex = (int)r.RowIndex.Value;
                    iGuessRowIndex = iRowIndex;
                }
                if (r.StyleIndex != null)
                {
                    index = (int)r.StyleIndex.Value;
                    // default is 0
                    if (index > 0 && index < listStyle.Count)
                    {
                        rp.StyleIndex = (uint)index;
                    }
                }
                if (r.Height != null)
                {
                    rp.HasHeight = true;
                    rp.Height = r.Height.Value;
                }
                if (r.Hidden != null && r.Hidden.Value) rp.Hidden = r.Hidden.Value;
                if (r.OutlineLevel != null && r.OutlineLevel.Value > 0) rp.OutlineLevel = r.OutlineLevel.Value;
                if (r.Collapsed != null && r.Collapsed.Value) rp.Collapsed = r.Collapsed.Value;
                if (r.ThickTop != null && r.ThickTop.Value) rp.ThickTop = r.ThickTop.Value;
                if (r.ThickBot != null && r.ThickBot.Value) rp.ThickBottom = r.ThickBot.Value;
                if (r.ShowPhonetic != null && r.ShowPhonetic.Value) rp.ShowPhonetic = r.ShowPhonetic.Value;

                if (slws.RowProperties.ContainsKey(iGuessRowIndex))
                {
                    slws.RowProperties[iGuessRowIndex] = rp;
                }
                else
                {
                    slws.RowProperties.Add(iGuessRowIndex, rp);
                }

                oxrRow = OpenXmlReader.Create(r);
                while (oxrRow.Read())
                {
                    if (oxrRow.ElementType == typeof(Cell))
                    {
                        ++iGuessColumnIndex;
                        c = (Cell)oxrRow.LoadCurrentElement();
                        slc = new SLCell();
                        slc.FromCell(c);

                        // 22 Oct 2016: This is reassign any existing InlineString to the
                        // shared string table.
                        if (c.DataType != null && c.DataType.Value == CellValues.InlineString && c.InlineString != null)
                        {
                            slc.DataType = CellValues.SharedString;
                            slc.NumericValue = this.DirectSaveToSharedStringTable(c.InlineString);
                        }

                        if (c.CellReference != null)
                        {
                            if (SLTool.FormatCellReferenceToRowColumnIndex(c.CellReference.Value, out iRowIndex, out iColumnIndex))
                            {
                                iGuessRowIndex = iRowIndex;
                                iGuessColumnIndex = iColumnIndex;
                                slws.CellWarehouse.SetValue(iGuessRowIndex, iGuessColumnIndex, slc);
                            }
                            else
                            {
                                slws.CellWarehouse.SetValue(iGuessRowIndex, iGuessColumnIndex, slc);
                            }
                        }
                        else
                        {
                            slws.CellWarehouse.SetValue(iGuessRowIndex, iGuessColumnIndex, slc);
                        }

                        // only if there's a cell formula
                        // and if the formula type is not null and is Shared
                        // and if there's a reference (list of cell ranges that links to this shared cell formula)
                        // and if there's a shared index.
                        // Open XML specs state that a shared cell formula must be Shared and there's a shared index.
                        // Cells with a reference will be the base cells, and cells without a reference will
                        // refer to those base cells.
                        // And the base cells must be uniquely identified by the shared index.
                        // Meaning if the reference property is not null, then the shared index must be unique.
                        // For example, you can't have a cell formula with
                        // ref="A1:B3" si=0
                        // and another cell formula with
                        // ref="F8:L10" si=0
                        // I'm only checking for well-formed spreadsheets, so if it comes through, I'm assuming
                        // the spreadsheet is ok. Like created by Excel would be good for a start.
                        if (c.CellFormula != null
                            && (c.CellFormula.FormulaType != null && c.CellFormula.FormulaType == CellFormulaValues.Shared)
                            && c.CellFormula.Reference != null
                            && c.CellFormula.SharedIndex != null)
                        {
                            if (!slws.SharedCellFormulas.ContainsKey(c.CellFormula.SharedIndex.Value))
                            {
                                scf = new SLSharedCellFormula();
                                scf.BaseCellRowIndex = iGuessRowIndex;
                                scf.BaseCellColumnIndex = iGuessColumnIndex;
                                scf.SharedIndex = c.CellFormula.SharedIndex;
                                scf.Reference = SLTool.TranslateReferenceToCellPointRangeList(c.CellFormula.Reference);
                                scf.FormulaText = c.CellFormula.Text;
                                slws.SharedCellFormulas.Add(c.CellFormula.SharedIndex.Value, scf);

                                // then we remove it. So the only place we note the cell formula is with
                                // the shared list. Don't worry, we put it back when we're writing the worksheet.
                                c.CellFormula = null;
                            }
                        }
                    }
                }
                oxrRow.Close();
                #endregion
            }
            else if (oxr.ElementType == typeof(SheetProtection))
            {
                SLSheetProtection sp = new SLSheetProtection();
                sp.FromSheetProtection((SheetProtection)oxr.LoadCurrentElement());
                slws.HasSheetProtection = true;
                slws.SheetProtection = sp.Clone();
            }
            else if (oxr.ElementType == typeof(AutoFilter))
            {
                SLAutoFilter af = new SLAutoFilter();
                af.FromAutoFilter((AutoFilter)oxr.LoadCurrentElement());
                slws.HasAutoFilter = true;
                slws.AutoFilter = af.Clone();
            }
            else if (oxr.ElementType == typeof(MergeCell))
            {
                mc = (MergeCell)oxr.LoadCurrentElement();
                slmc = new SLMergeCell();
                slmc.FromMergeCell(mc);
                if (slmc.IsValid) slws.MergeCells.Add(slmc);
            }
            else if (oxr.ElementType == typeof(ConditionalFormatting))
            {
                condformat = new SLConditionalFormatting();
                condformat.FromConditionalFormatting((ConditionalFormatting)oxr.LoadCurrentElement());
                slws.ConditionalFormattings.Add(condformat);
            }
            else if (oxr.ElementType == typeof(DataValidations))
            {
                DataValidations dvs = (DataValidations)oxr.LoadCurrentElement();
                if (dvs.DisablePrompts != null) slws.DataValidationDisablePrompts = dvs.DisablePrompts.Value;
                if (dvs.XWindow != null) slws.DataValidationXWindow = dvs.XWindow.Value;
                if (dvs.YWindow != null) slws.DataValidationYWindow = dvs.YWindow.Value;

                using (OpenXmlReader oxrDataValidation = OpenXmlReader.Create(dvs))
                {
                    SLDataValidation dv;
                    while (oxrDataValidation.Read())
                    {
                        if (oxrDataValidation.ElementType == typeof(DataValidation))
                        {
                            dv = new SLDataValidation();
                            dv.FromDataValidation((DataValidation)oxrDataValidation.LoadCurrentElement());
                            slws.DataValidations.Add(dv);
                        }
                    }
                }
            }
            else if (oxr.ElementType == typeof(Hyperlink))
            {
                hl = new SLHyperlink();
                hl.FromHyperlink((Hyperlink)oxr.LoadCurrentElement());
                slws.Hyperlinks.Add(hl);
            }
            else if (oxr.ElementType == typeof(PrintOptions))
            {
                PrintOptions po = (PrintOptions)oxr.LoadCurrentElement();
                if (po.HorizontalCentered != null) slws.PageSettings.PrintHorizontalCentered = po.HorizontalCentered.Value;
                if (po.VerticalCentered != null) slws.PageSettings.PrintVerticalCentered = po.VerticalCentered.Value;
                if (po.Headings != null) slws.PageSettings.PrintHeadings = po.Headings.Value;
                if (po.GridLines != null) slws.PageSettings.PrintGridLines = po.GridLines.Value;
                if (po.GridLinesSet != null) slws.PageSettings.PrintGridLinesSet = po.GridLinesSet.Value;
            }
            else if (oxr.ElementType == typeof(PageMargins))
            {
                PageMargins pm = (PageMargins)oxr.LoadCurrentElement();
                if (pm.Left != null) slws.PageSettings.LeftMargin = pm.Left.Value;
                if (pm.Right != null) slws.PageSettings.RightMargin = pm.Right.Value;
                if (pm.Top != null) slws.PageSettings.TopMargin = pm.Top.Value;
                if (pm.Bottom != null) slws.PageSettings.BottomMargin = pm.Bottom.Value;
                if (pm.Header != null) slws.PageSettings.HeaderMargin = pm.Header.Value;
                if (pm.Footer != null) slws.PageSettings.FooterMargin = pm.Footer.Value;
            }
            else if (oxr.ElementType == typeof(PageSetup))
            {
                PageSetup ps = (PageSetup)oxr.LoadCurrentElement();
                // consider setting to 1 if not one of the "valid" paper sizes?
                if (ps.PaperSize != null) slws.PageSettings.PaperSize = (SLPaperSizeValues)ps.PaperSize.Value;
                if (ps.Scale != null) slws.PageSettings.iScale = ps.Scale.Value;
                if (ps.FirstPageNumber != null) slws.PageSettings.FirstPageNumber = ps.FirstPageNumber.Value;
                if (ps.FitToWidth != null) slws.PageSettings.iFitToWidth = ps.FitToWidth.Value;
                if (ps.FitToHeight != null) slws.PageSettings.iFitToHeight = ps.FitToHeight.Value;
                if (ps.PageOrder != null) slws.PageSettings.PageOrder = ps.PageOrder.Value;
                if (ps.Orientation != null) slws.PageSettings.Orientation = ps.Orientation.Value;
                if (ps.UsePrinterDefaults != null) slws.PageSettings.UsePrinterDefaults = ps.UsePrinterDefaults.Value;
                if (ps.BlackAndWhite != null) slws.PageSettings.BlackAndWhite = ps.BlackAndWhite.Value;
                if (ps.Draft != null) slws.PageSettings.Draft = ps.Draft.Value;
                if (ps.CellComments != null) slws.PageSettings.CellComments = ps.CellComments.Value;
                if (ps.Errors != null) slws.PageSettings.Errors = ps.Errors.Value;
                if (ps.HorizontalDpi != null) slws.PageSettings.HorizontalDpi = ps.HorizontalDpi.Value;
                if (ps.VerticalDpi != null) slws.PageSettings.VerticalDpi = ps.VerticalDpi.Value;
                if (ps.Copies != null) slws.PageSettings.Copies = ps.Copies.Value;
            }
            else if (oxr.ElementType == typeof(HeaderFooter))
            {
                HeaderFooter hf = (HeaderFooter)oxr.LoadCurrentElement();
                if (hf.OddHeader != null) slws.PageSettings.OddHeaderText = hf.OddHeader.Text;
                if (hf.OddFooter != null) slws.PageSettings.OddFooterText = hf.OddFooter.Text;
                if (hf.EvenHeader != null) slws.PageSettings.EvenHeaderText = hf.EvenHeader.Text;
                if (hf.EvenFooter != null) slws.PageSettings.EvenFooterText = hf.EvenFooter.Text;
                if (hf.FirstHeader != null) slws.PageSettings.FirstHeaderText = hf.FirstHeader.Text;
                if (hf.FirstFooter != null) slws.PageSettings.FirstFooterText = hf.FirstFooter.Text;
                if (hf.DifferentOddEven != null) slws.PageSettings.DifferentOddEvenPages = hf.DifferentOddEven.Value;
                if (hf.DifferentFirst != null) slws.PageSettings.DifferentFirstPage = hf.DifferentFirst.Value;
                if (hf.ScaleWithDoc != null) slws.PageSettings.ScaleWithDocument = hf.ScaleWithDoc.Value;
                if (hf.AlignWithMargins != null) slws.PageSettings.AlignWithMargins = hf.AlignWithMargins.Value;
            }
            else if (oxr.ElementType == typeof(RowBreaks))
            {
                SLBreak b;
                uint rowbkindex;
                using (OpenXmlReader oxrRowBreaks = OpenXmlReader.Create((RowBreaks)oxr.LoadCurrentElement()))
                {
                    while (oxrRowBreaks.Read())
                    {
                        if (oxrRowBreaks.ElementType == typeof(Break))
                        {
                            b = new SLBreak();
                            b.FromBreak((Break)oxrRowBreaks.LoadCurrentElement());
                            rowbkindex = b.Id;
                            slws.RowBreaks[(int)rowbkindex] = b;
                        }
                    }
                }
            }
            else if (oxr.ElementType == typeof(ColumnBreaks))
            {
                SLBreak b;
                uint colbkindex;
                using (OpenXmlReader oxrColBreaks = OpenXmlReader.Create((ColumnBreaks)oxr.LoadCurrentElement()))
                {
                    while (oxrColBreaks.Read())
                    {
                        if (oxrColBreaks.ElementType == typeof(Break))
                        {
                            b = new SLBreak();
                            b.FromBreak((Break)oxrColBreaks.LoadCurrentElement());
                            colbkindex = b.Id;
                            slws.ColumnBreaks[(int)colbkindex] = b;
                        }
                    }
                }
            }
            else if (oxr.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Drawing))
            {
                DocumentFormat.OpenXml.Spreadsheet.Drawing drawing = (DocumentFormat.OpenXml.Spreadsheet.Drawing)oxr.LoadCurrentElement();
                slws.DrawingId = drawing.Id;
                if (wsp.DrawingsPart != null)
                {
                    Xdr.NonVisualDrawingProperties nvdp;
                    uint iUniqueId = 1;
                    using (OpenXmlReader oxrDrawing = OpenXmlReader.Create(wsp.DrawingsPart.WorksheetDrawing))
                    {
                        while (oxrDrawing.Read())
                        {
                            if (oxrDrawing.ElementType == typeof(Xdr.NonVisualDrawingProperties))
                            {
                                nvdp = (Xdr.NonVisualDrawingProperties)oxrDrawing.LoadCurrentElement();
                                if (nvdp.Id != null && nvdp.Id.Value > iUniqueId)
                                {
                                    iUniqueId = nvdp.Id.Value;
                                }
                            }
                        }
                    }
                    slws.NextWorksheetDrawingId = iUniqueId + 1;
                }
            }
            else if (oxr.ElementType == typeof(Picture))
            {
                Picture pic = (Picture)oxr.LoadCurrentElement();
                slws.BackgroundPictureId = pic.Id;
                slws.BackgroundPictureDataIsInFile = null;
            }
            else if (oxr.ElementType == typeof(WorksheetExtensionList))
            {
                WorksheetExtensionList wsextlist = (WorksheetExtensionList)oxr.LoadCurrentElement();

                SLConditionalFormatting2010 cf2010;
                X14.SparklineGroup sparkgrp;
                SLSparklineGroup spkgrp;

                using (OpenXmlReader oxrext = OpenXmlReader.Create(wsextlist))
                {
                    while (oxrext.Read())
                    {
                        if (oxrext.ElementType == typeof(X14.ConditionalFormatting))
                        {
                            cf2010 = new SLConditionalFormatting2010();
                            cf2010.FromConditionalFormatting((X14.ConditionalFormatting)oxrext.LoadCurrentElement());
                            slws.ConditionalFormattings2010.Add(cf2010.Clone());
                        }
                        else if (oxrext.ElementType == typeof(X14.SparklineGroup))
                        {
                            sparkgrp = (X14.SparklineGroup)oxrext.LoadCurrentElement();
                            spkgrp = new SLSparklineGroup(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);
                            spkgrp.FromSparklineGroup(sparkgrp);
                            slws.SparklineGroups.Add(spkgrp.Clone());
                        }
                    }
                }
            }
        }
        oxr.Dispose();

        if (wsp.TableDefinitionParts != null)
        {
            SLTable t;
            foreach (TableDefinitionPart tdp in wsp.TableDefinitionParts)
            {
                t = new SLTable();
                t.FromTable(tdp.Table);
                t.RelationshipID = wsp.GetIdOfPart(tdp);
                t.IsNewTable = false;
                slws.Tables.Add(t);
            }
        }

        // TODO: as at 17 Jan 2014, I have forgotten what I planned to do here. Go me...
        // I think there's an order and priority to the conditional formats, and it depends
        // on whether it's normal conditional formats or the new 2010 version.
        // I must've done some research, but I can't find the results already... poor future Vincent...

        // determine priority and order of conditional formats
        //if (slws.ConditionalFormattings.Count == 0 && slws.ConditionalFormattings2010.Count == 0)
        //{
        //    // this is probably the most common, so put it as the first condition (pun, haha)
        //    // for optimisation. Aaaannnddd, we don't have to do anything here. Awesome, right?
        //}
        //else if (slws.ConditionalFormattings.Count > 0 && slws.ConditionalFormattings2010.Count == 0)
        //{
        //    for (int i = 0; i < slws.ConditionalFormattings.Count; ++i)
        //    {
        //        slws.ConditionalFormatOrder.Add(new SLConditionalFormatOrder()
        //        {
        //            Version = SLConditionalFormatVersionValue.Office
        //        });
        //    }
        //}
        //else if (slws.ConditionalFormattings.Count == 0 && slws.ConditionalFormattings2010.Count > 0)
        //{   
        //}
        //else if (slws.ConditionalFormattings.Count > 0 && slws.ConditionalFormattings2010.Count > 0)
        //{
        //}
        // no else because all conditions already in one of the if conditions.
    }

    private void WriteSelectedWorksheet()
    {
        // split into writing for existing worksheet and for new worksheet

        this.CleanUpReallyEmptyCells();

        int i = 0, j = 0, k = 0;
        bool bFound = false;
        OpenXmlElement oxe;

        SLColumnProperties cp;
        SLRowProperties rp;
        byte byMaxOutline;
        List<int> listintkeys;

        // remove empty rows/columns plus getting the maximum outline levels at the same time.

        byMaxOutline = 0;
        listintkeys = slws.RowProperties.Keys.ToList<int>();
        foreach (int key in listintkeys)
        {
            rp = slws.RowProperties[key];
            if (rp.IsEmpty) slws.RowProperties.Remove(key);
            else if (rp.OutlineLevel > byMaxOutline) byMaxOutline = rp.OutlineLevel;
        }
        if (byMaxOutline > 0) slws.SheetFormatProperties.OutlineLevelRow = byMaxOutline;

        byMaxOutline = 0;
        listintkeys = slws.ColumnProperties.Keys.ToList<int>();
        foreach (int key in listintkeys)
        {
            cp = slws.ColumnProperties[key];
            if (cp.IsEmpty) slws.ColumnProperties.Remove(key);
            else if (cp.OutlineLevel > byMaxOutline) byMaxOutline = cp.OutlineLevel;
        }
        if (byMaxOutline > 0) slws.SheetFormatProperties.OutlineLevelColumn = byMaxOutline;

        SLCell c;
        int iStartRowIndex = SLConstants.RowLimit + 1;
        int iStartColumnIndex = SLConstants.ColumnLimit + 1;
        int iEndRowIndex = -1;
        int iEndColumnIndex = -1;
        // in case somehow the base cell doesn't exist. I mean, there's a lot going on for shared formulas...
        foreach (var scfiter in slws.SharedCellFormulas)
        {
            for (k = scfiter.Value.Reference.Count - 1; k >= 0; --k)
            {
                iStartRowIndex = scfiter.Value.Reference[k].StartRowIndex;
                iEndRowIndex = scfiter.Value.Reference[k].EndRowIndex;
                iStartColumnIndex = scfiter.Value.Reference[k].StartColumnIndex;
                iEndColumnIndex = scfiter.Value.Reference[k].EndColumnIndex;
                for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                {
                    for (j = iStartColumnIndex; j <= iEndColumnIndex; ++j)
                    {
                        if (i == scfiter.Value.BaseCellRowIndex && j == scfiter.Value.BaseCellColumnIndex)
                        {
                            // is base cell

                            if (slws.CellWarehouse.Exists(i, j))
                            {
                                c = slws.CellWarehouse.Cells[i][j].Clone();
                                if (c.CellFormula == null) c.CellFormula = new SLCellFormula();
                                // need to set this? Whatever, just set and see what happens..
                                c.CellFormula.SharedIndex = scfiter.Value.SharedIndex;
                                c.CellFormula.Reference.Clear();
                                foreach (SLCellPointRange scfptrange in scfiter.Value.Reference)
                                {
                                    c.CellFormula.Reference.Add(new SLCellPointRange(
                                        scfptrange.StartRowIndex,
                                        scfptrange.StartColumnIndex,
                                        scfptrange.EndRowIndex,
                                        scfptrange.EndColumnIndex));
                                }
                                c.CellFormula.FormulaText = scfiter.Value.FormulaText;
                                c.CellText = string.Empty;
                                slws.CellWarehouse.SetValue(i, j, c);
                            }
                            else
                            {
                                // there must (??) be a base cell existing in the worksheet
                                // (otherwise where to find the formula?)
                                c = new SLCell();
                                c.DataType = CellValues.Number;
                                c.CellFormula = new SLCellFormula();
                                c.CellFormula.SharedIndex = scfiter.Value.SharedIndex;
                                foreach (SLCellPointRange scfptrange in scfiter.Value.Reference)
                                {
                                    c.CellFormula.Reference.Add(new SLCellPointRange(
                                        scfptrange.StartRowIndex,
                                        scfptrange.StartColumnIndex,
                                        scfptrange.EndRowIndex,
                                        scfptrange.EndColumnIndex));
                                }
                                c.CellFormula.FormulaText = scfiter.Value.FormulaText;
                                c.CellText = string.Empty;
                                slws.CellWarehouse.SetValue(i, j, c);
                            }
                        }
                        else
                        {
                            // is not base cell

                            if (slws.CellWarehouse.Exists(i, j))
                            {
                                c = slws.CellWarehouse.Cells[i][j].Clone();
                                // if not null, then we assume it has an actual directly set formula
                                // (or Array or DataTable type, but I don't know what to do with those...)
                                if (c.CellFormula == null)
                                {
                                    c.DataType = CellValues.Number;
                                    c.CellFormula = new SLCellFormula();
                                    c.CellFormula.FormulaType = CellFormulaValues.Shared;
                                    c.CellFormula.SharedIndex = scfiter.Value.SharedIndex;
                                    slws.CellWarehouse.SetValue(i, j, c);
                                }
                            }
                        }
                    }
                }
            }
        }

        List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
        List<int> listColumnKeys;

        foreach (int rpkey in slws.RowProperties.Keys.ToList<int>())
        {
            if (!listRowKeys.Contains(rpkey))
            {
                listRowKeys.Add(rpkey);
            }
        }

        // this now contains every row index that's either in the list of row properties
        // or in the list of cells.
        listRowKeys.Sort();
        // also note that because the row indices of the RowProperties is included,
        // the code below has to check for the index before accessing the cell dictionary.

        List<int> listColumnIndex = slws.ColumnProperties.Keys.ToList<int>();
        listColumnIndex.Sort();

        // reusing the index variables for calculating the dimension property
        iStartRowIndex = SLConstants.RowLimit + 1;
        iStartColumnIndex = SLConstants.ColumnLimit + 1;
        iEndRowIndex = -1;
        iEndColumnIndex = -1;

        if (listRowKeys.Count > 0 || listColumnIndex.Count > 0 || slws.MergeCells.Count > 0)
        {
            foreach (int rowkey in listRowKeys)
            {
                if (rowkey < iStartRowIndex) iStartRowIndex = rowkey;
                if (rowkey > iEndRowIndex) iEndRowIndex = rowkey;

                if (slws.CellWarehouse.Cells.ContainsKey(rowkey))
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        if (colkey < iStartColumnIndex) iStartColumnIndex = colkey;
                        if (colkey > iEndColumnIndex) iEndColumnIndex = colkey;
                    }
                }
            }

            foreach (SLMergeCell mc in slws.MergeCells)
            {
                if (mc.StartRowIndex < iStartRowIndex) iStartRowIndex = mc.StartRowIndex;
                if (mc.StartColumnIndex < iStartColumnIndex) iStartColumnIndex = mc.StartColumnIndex;
                if (mc.EndRowIndex > iEndRowIndex) iEndRowIndex = mc.EndRowIndex;
                if (mc.EndColumnIndex > iEndColumnIndex) iEndColumnIndex = mc.EndColumnIndex;
            }

            // need to do for hyperlinks?
            //foreach (SLHyperlink hl in slws.Hyperlinks)
            //{
            //    if (hl.Reference.StartRowIndex < iDimensionStartRowIndex) iDimensionStartRowIndex = hl.Reference.StartRowIndex;
            //    if (hl.Reference.StartColumnIndex < iDimensionStartColumnIndex) iDimensionStartColumnIndex = hl.Reference.StartColumnIndex;
            //    if (hl.Reference.EndRowIndex > iDimensionEndRowIndex) iDimensionEndRowIndex = hl.Reference.EndRowIndex;
            //    if (hl.Reference.EndColumnIndex > iDimensionEndColumnIndex) iDimensionEndColumnIndex = hl.Reference.EndColumnIndex;
            //}
        }

        string sDimensionCellRange = string.Empty;
        if (iStartRowIndex > SLConstants.RowLimit) iStartRowIndex = 1;
        if (iStartColumnIndex > SLConstants.ColumnLimit) iStartColumnIndex = 1;
        if (iEndRowIndex < 1) iEndRowIndex = 1;
        if (iEndColumnIndex < 1) iEndColumnIndex = 1;
        if (iStartRowIndex == iEndRowIndex && iStartColumnIndex == iEndColumnIndex)
        {
            sDimensionCellRange = SLTool.ToCellReference(iStartRowIndex, iStartColumnIndex);
        }
        else
        {
            sDimensionCellRange = string.Format("{0}:{1}", SLTool.ToCellReference(iStartRowIndex, iStartColumnIndex), SLTool.ToCellReference(iEndRowIndex, iEndColumnIndex));
        }

        Row r;

        if (!IsNewWorksheet)
        {
            // Need to check?
            //if (string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID)) return;

            WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);

            if (slws.ForceCustomRowColumnDimensionsSplitting)
            {
                slws.ToggleCustomRowColumnDimension(true);
            }

            if (slws.PageSettings.HasSheetProperties)
            {
                wsp.Worksheet.SheetProperties = slws.PageSettings.SheetProperties.ToSheetProperties();
            }
            else
            {
                wsp.Worksheet.SheetProperties = null;
            }

            wsp.Worksheet.SheetDimension = new SheetDimension() { Reference = sDimensionCellRange };

            if (slws.SheetViews.Count > 0)
            {
                wsp.Worksheet.SheetViews = new SheetViews();
                foreach (SLSheetView sv in slws.SheetViews)
                {
                    wsp.Worksheet.SheetViews.Append(sv.ToSheetView());
                }
            }
            else
            {
                wsp.Worksheet.SheetViews = null;
            }

            wsp.Worksheet.SheetFormatProperties = slws.SheetFormatProperties.ToSheetFormatProperties();

            #region Filling Columns
            if (wsp.Worksheet.Elements<Columns>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<Columns>();
            }

            if (slws.ColumnProperties.Count > 0)
            {
                Columns cols = new Columns();
                Column col;

                int iPreviousColumnIndex = listColumnIndex[0];
                int iCurrentColumnIndex = iPreviousColumnIndex;
                string sCollectiveColumnData = string.Empty;
                string sCurrentColumnData = string.Empty;
                int colmin, colmax;
                colmin = colmax = iCurrentColumnIndex;
                cp = slws.ColumnProperties[iCurrentColumnIndex];
                sCollectiveColumnData = cp.ToHash();

                col = new Column();
                col.Min = (uint)colmin;
                col.Max = (uint)colmax;
                if (cp.HasWidth)
                {
                    col.Width = cp.Width;
                    col.CustomWidth = true;
                }
                else
                {
                    col.Width = slws.SheetFormatProperties.DefaultColumnWidth;
                }
                if (cp.StyleIndex > 0) col.Style = cp.StyleIndex;
                if (cp.Hidden) col.Hidden = cp.Hidden;
                if (cp.BestFit) col.BestFit = cp.BestFit;
                if (cp.Phonetic) col.Phonetic = cp.Phonetic;
                if (cp.OutlineLevel > 0) col.OutlineLevel = cp.OutlineLevel;
                if (cp.Collapsed) col.Collapsed = cp.Collapsed;

                for (i = 1; i < listColumnIndex.Count; ++i)
                {
                    iPreviousColumnIndex = iCurrentColumnIndex;
                    iCurrentColumnIndex = listColumnIndex[i];
                    cp = slws.ColumnProperties[iCurrentColumnIndex];
                    sCurrentColumnData = cp.ToHash();

                    if ((iCurrentColumnIndex != (iPreviousColumnIndex + 1)) || (sCollectiveColumnData != sCurrentColumnData))
                    {
                        col.Max = (uint)colmax;
                        cols.Append(col);

                        colmin = iCurrentColumnIndex;
                        colmax = iCurrentColumnIndex;
                        sCollectiveColumnData = sCurrentColumnData;

                        col = new Column();
                        col.Min = (uint)colmin;
                        col.Max = (uint)colmax;
                        if (cp.HasWidth)
                        {
                            col.Width = cp.Width;
                            col.CustomWidth = true;
                        }
                        else
                        {
                            col.Width = slws.SheetFormatProperties.DefaultColumnWidth;
                        }
                        if (cp.StyleIndex > 0) col.Style = cp.StyleIndex;
                        if (cp.Hidden) col.Hidden = cp.Hidden;
                        if (cp.BestFit) col.BestFit = cp.BestFit;
                        if (cp.Phonetic) col.Phonetic = cp.Phonetic;
                        if (cp.OutlineLevel > 0) col.OutlineLevel = cp.OutlineLevel;
                        if (cp.Collapsed) col.Collapsed = cp.Collapsed;
                    }
                    else
                    {
                        colmax = iCurrentColumnIndex;
                    }
                }

                // there's always a "leftover" column
                col.Max = (uint)colmax;
                cols.Append(col);

                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    if (child is SheetProperties || child is SheetDimension || child is SheetViews || child is SheetFormatProperties)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(cols, oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(cols);
                }
            }
            #endregion

            SheetData sd = new SheetData();

            foreach (int rowkey in listRowKeys)
            {
                if (slws.RowProperties.ContainsKey(rowkey))
                {
                    r = slws.RowProperties[rowkey].ToRow();
                    r.RowIndex = (uint)rowkey;
                }
                else
                {
                    r = new Row();
                    r.RowIndex = (uint)rowkey;
                }

                if (slws.CellWarehouse.Cells.ContainsKey(rowkey))
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    listColumnKeys.Sort();
                    foreach (int colkey in listColumnKeys)
                    {
                        c = slws.CellWarehouse.Cells[rowkey][colkey].Clone();
                        r.Append(c.ToCell(SLTool.ToCellReference(rowkey, colkey)));
                    }
                }
                sd.Append(r);
            }

            wsp.Worksheet.RemoveAllChildren<SheetData>();

            bFound = false;
            oxe = wsp.Worksheet.FirstChild;
            foreach (var child in wsp.Worksheet.ChildElements)
            {
                if (child is SheetProperties || child is SheetDimension || child is SheetViews || child is SheetFormatProperties || child is Columns)
                {
                    oxe = child;
                    bFound = true;
                }
            }

            if (bFound)
            {
                wsp.Worksheet.InsertAfter(sd, oxe);
            }
            else
            {
                wsp.Worksheet.PrependChild(sd);
            }

            #region Sheet protection
            if (wsp.Worksheet.Elements<SheetProtection>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<SheetProtection>();
            }

            if (slws.HasSheetProtection)
            {
                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(slws.SheetProtection.ToSheetProtection(), oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(slws.SheetProtection.ToSheetProtection());
                }
            }
            #endregion

            #region AutoFilter
            if (wsp.Worksheet.Elements<AutoFilter>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<AutoFilter>();
            }

            if (slws.HasAutoFilter)
            {
                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(slws.AutoFilter.ToAutoFilter(), oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(slws.AutoFilter.ToAutoFilter());
                }
            }
            #endregion

            #region Filling merge cells
            if (wsp.Worksheet.Elements<MergeCells>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<MergeCells>();
            }

            if (slws.MergeCells.Count > 0)
            {
                MergeCells mcs = new MergeCells() { Count = (uint)slws.MergeCells.Count };
                for (i = 0; i < slws.MergeCells.Count; ++i)
                {
                    mcs.Append(slws.MergeCells[i].ToMergeCell());
                }

                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(mcs, oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(mcs);
                }
            }
            #endregion

            #region Conditional Formatting
            if (wsp.Worksheet.Elements<ConditionalFormatting>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<ConditionalFormatting>();
            }

            if (slws.ConditionalFormattings.Count > 0)
            {
                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    for (i = slws.ConditionalFormattings.Count - 1; i >= 0; --i)
                    {
                        wsp.Worksheet.InsertAfter(slws.ConditionalFormattings[i].ToConditionalFormatting(), oxe);
                    }
                }
                else
                {
                    for (i = slws.ConditionalFormattings.Count - 1; i >= 0; --i)
                    {
                        wsp.Worksheet.PrependChild(slws.ConditionalFormattings[i].ToConditionalFormatting());
                    }
                }
            }
            #endregion

            #region DataValidations
            if (wsp.Worksheet.Elements<DataValidations>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<DataValidations>();
            }

            if (slws.DataValidations.Count > 0)
            {
                DataValidations dvs = new DataValidations();
                if (slws.DataValidationDisablePrompts) dvs.DisablePrompts = slws.DataValidationDisablePrompts;
                if (slws.DataValidationXWindow != null) dvs.XWindow = slws.DataValidationXWindow.Value;
                if (slws.DataValidationYWindow != null) dvs.YWindow = slws.DataValidationYWindow.Value;
                dvs.Count = (uint)slws.DataValidations.Count;

                foreach (SLDataValidation dv in slws.DataValidations)
                {
                    dvs.Append(dv.ToDataValidation());
                }

                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(dvs, oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(dvs);
                }
            }
            #endregion

            #region Hyperlinks
            if (wsp.Worksheet.Elements<Hyperlinks>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<Hyperlinks>();
            }

            if (slws.Hyperlinks.Count > 0)
            {
                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                        || child is DataValidations)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                Hyperlinks hls = new Hyperlinks();
                HyperlinkRelationship hlrel;
                foreach (SLHyperlink hl in slws.Hyperlinks)
                {
                    if (hl.IsExternal && hl.IsNew)
                    {
                        hlrel = wsp.AddHyperlinkRelationship(new Uri(hl.HyperlinkUri, hl.HyperlinkUriKind), true);
                        hl.Id = hlrel.Id;
                    }
                    hls.Append(hl.ToHyperlink());
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(hls, oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(hls);
                }
            }
            #endregion

            #region PrintOptions
            if (wsp.Worksheet.Elements<PrintOptions>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<PrintOptions>();
            }

            if (slws.PageSettings.HasPrintOptions)
            {
                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                        || child is DataValidations || child is Hyperlinks)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(slws.PageSettings.ExportPrintOptions(), oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(slws.PageSettings.ExportPrintOptions());
                }
            }
            #endregion

            #region PageMargins
            if (wsp.Worksheet.Elements<PageMargins>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<PageMargins>();
            }

            if (slws.PageSettings.HasPageMargins)
            {
                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                        || child is DataValidations || child is Hyperlinks || child is PrintOptions)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(slws.PageSettings.ExportPageMargins(), oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(slws.PageSettings.ExportPageMargins());
                }
            }
            #endregion

            #region PageSetup
            if (wsp.Worksheet.Elements<PageSetup>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<PageSetup>();
            }

            if (slws.PageSettings.HasPageSetup)
            {
                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                        || child is DataValidations || child is Hyperlinks || child is PrintOptions
                        || child is PageMargins)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(slws.PageSettings.ExportPageSetup(), oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(slws.PageSettings.ExportPageSetup());
                }
            }
            #endregion

            #region HeaderFooter
            if (wsp.Worksheet.Elements<HeaderFooter>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<HeaderFooter>();
            }

            if (slws.PageSettings.HasHeaderFooter)
            {
                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                        || child is DataValidations || child is Hyperlinks || child is PrintOptions
                        || child is PageMargins || child is PageSetup)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(slws.PageSettings.ExportHeaderFooter(), oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(slws.PageSettings.ExportHeaderFooter());
                }
            }
            #endregion

            #region RowBreaks
            if (wsp.Worksheet.Elements<RowBreaks>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<RowBreaks>();
            }

            if (slws.RowBreaks.Count > 0)
            {
                List<int> bkkeys = slws.RowBreaks.Keys.ToList<int>();
                bkkeys.Sort();

                RowBreaks rowbk = new RowBreaks();
                int bkmancount = 0;
                foreach (int bkindex in bkkeys)
                {
                    if (slws.RowBreaks[bkindex].ManualPageBreak) ++bkmancount;
                    rowbk.Append(slws.RowBreaks[bkindex].ToBreak());
                }
                rowbk.Count = (uint)slws.RowBreaks.Count;
                rowbk.ManualBreakCount = (uint)bkmancount;

                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                        || child is DataValidations || child is Hyperlinks || child is PrintOptions
                        || child is PageMargins || child is PageSetup || child is HeaderFooter)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(rowbk, oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(rowbk);
                }
            }
            #endregion

            #region ColumnBreaks
            if (wsp.Worksheet.Elements<ColumnBreaks>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<ColumnBreaks>();
            }

            if (slws.ColumnBreaks.Count > 0)
            {
                List<int> bkkeys = slws.ColumnBreaks.Keys.ToList<int>();
                bkkeys.Sort();

                ColumnBreaks colbk = new ColumnBreaks();
                int bkmancount = 0;
                foreach (int bkindex in bkkeys)
                {
                    if (slws.ColumnBreaks[bkindex].ManualPageBreak) ++bkmancount;
                    colbk.Append(slws.ColumnBreaks[bkindex].ToBreak());
                }
                colbk.Count = (uint)slws.ColumnBreaks.Count;
                colbk.ManualBreakCount = (uint)bkmancount;

                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                        || child is DataValidations || child is Hyperlinks || child is PrintOptions
                        || child is PageMargins || child is PageSetup || child is HeaderFooter
                        || child is RowBreaks)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(colbk, oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(colbk);
                }
            }
            #endregion

            #region Drawing
            // these are "new" charts and pictures added
            if (slws.Charts.Count > 0 || slws.Pictures.Count > 0)
            {
                // if the length > 0, then we assume there's already an existing DrawingsPart
                if (slws.DrawingId.Length > 0)
                {
                    WriteImageParts(wsp.DrawingsPart);
                }
                else
                {
                    DrawingsPart dp = wsp.AddNewPart<DrawingsPart>();
                    dp.WorksheetDrawing = new Xdr.WorksheetDrawing();
                    dp.WorksheetDrawing.AddNamespaceDeclaration("xdr", SLConstants.NamespaceXdr);
                    dp.WorksheetDrawing.AddNamespaceDeclaration("a", SLConstants.NamespaceA);

                    DocumentFormat.OpenXml.Spreadsheet.Drawing drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing();
                    drawing.Id = wsp.GetIdOfPart(dp);

                    WriteImageParts(dp);

                    // NOTE: SmartTags is deprecated in Open XML SDK 2.5, so have to remove
                    // from check below?

                    bFound = false;
                    oxe = wsp.Worksheet.FirstChild;
                    foreach (var child in wsp.Worksheet.ChildElements)
                    {
                        // start with SheetData because it's a required child element
                        if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                            || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                            || child is SortState || child is DataConsolidate || child is CustomSheetViews
                            || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                            || child is DataValidations || child is Hyperlinks || child is PrintOptions
                            || child is PageMargins || child is PageSetup || child is HeaderFooter
                            || child is RowBreaks || child is ColumnBreaks || child is CustomProperties
                            || child is CellWatches || child is IgnoredErrors)
                        //|| child is CellWatches || child is IgnoredErrors || child is SmartTags)
                        // SmartTags is deprecated...
                        {
                            oxe = child;
                            bFound = true;
                        }
                    }

                    if (bFound)
                    {
                        wsp.Worksheet.InsertAfter(drawing, oxe);
                    }
                    else
                    {
                        wsp.Worksheet.PrependChild(drawing);
                    }
                }
            }
            #endregion

            #region LegacyDrawing
            if (slws.Comments.Count > 0)
            {
                // we're going to do this only if there are no comments and VML already
                if (wsp.WorksheetCommentsPart == null
                    && wsp.Worksheet.Elements<LegacyDrawing>().Count() == 0)
                {
                    WorksheetCommentsPart wcp = wsp.AddNewPart<WorksheetCommentsPart>();
                    VmlDrawingPart vdp = wsp.AddNewPart<VmlDrawingPart>();
                    WriteCommentPart(wcp, vdp);

                    LegacyDrawing ldrawing = new LegacyDrawing();
                    ldrawing.Id = wsp.GetIdOfPart(vdp);

                    // NOTE: SmartTags is deprecated in Open XML SDK 2.5, so have to remove
                    // from check below?

                    bFound = false;
                    oxe = wsp.Worksheet.FirstChild;
                    foreach (var child in wsp.Worksheet.ChildElements)
                    {
                        // start with SheetData because it's a required child element
                        if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                            || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                            || child is SortState || child is DataConsolidate || child is CustomSheetViews
                            || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                            || child is DataValidations || child is Hyperlinks || child is PrintOptions
                            || child is PageMargins || child is PageSetup || child is HeaderFooter
                            || child is RowBreaks || child is ColumnBreaks || child is CustomProperties
                            || child is CellWatches || child is IgnoredErrors
                            //|| child is CellWatches || child is IgnoredErrors || child is SmartTags
                            || child is DocumentFormat.OpenXml.Spreadsheet.Drawing)
                        {
                            oxe = child;
                            bFound = true;
                        }
                    }

                    if (bFound)
                    {
                        wsp.Worksheet.InsertAfter(ldrawing, oxe);
                    }
                    else
                    {
                        wsp.Worksheet.PrependChild(ldrawing);
                    }
                }
            }
            #endregion

            #region Picture
            if (wsp.Worksheet.Elements<Picture>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<Picture>();
            }

            if (slws.BackgroundPictureId.Length > 0 || slws.BackgroundPictureDataIsInFile != null)
            {
                Picture pic = new Picture();
                if (slws.BackgroundPictureId.Length > 0)
                {
                    pic.Id = slws.BackgroundPictureId;
                }
                else if (slws.BackgroundPictureDataIsInFile != null)
                {
                    ImagePart imgp = wsp.AddImagePart(slws.BackgroundPictureImagePartType);
                    if (slws.BackgroundPictureDataIsInFile.Value)
                    {
                        using (FileStream fs = new FileStream(slws.BackgroundPictureFileName, FileMode.Open))
                        {
                            imgp.FeedData(fs);
                        }
                    }
                    else
                    {
                        using (MemoryStream ms = new MemoryStream(slws.BackgroundPictureByteData))
                        {
                            imgp.FeedData(ms);
                        }
                    }
                    pic.Id = wsp.GetIdOfPart(imgp);
                }

                // NOTE: SmartTags is deprecated in Open XML SDK 2.5, so have to remove
                // from check below?

                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                        || child is DataValidations || child is Hyperlinks || child is PrintOptions
                        || child is PageMargins || child is PageSetup || child is HeaderFooter
                        || child is RowBreaks || child is ColumnBreaks || child is CustomProperties
                        || child is CellWatches || child is IgnoredErrors
                        //|| child is CellWatches || child is IgnoredErrors || child is SmartTags
                        || child is DocumentFormat.OpenXml.Spreadsheet.Drawing
                        || child is LegacyDrawing || child is LegacyDrawingHeaderFooter)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(pic, oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(pic);
                }
            }
            #endregion

            #region Tables
            if (wsp.Worksheet.Elements<TableParts>().Count() > 0)
            {
                wsp.Worksheet.RemoveAllChildren<TableParts>();
            }

            if (slws.Tables.Count > 0)
            {
                TableParts tps = new TableParts() { Count = (uint)slws.Tables.Count };
                TableDefinitionPart tdp;
                string sRelID = string.Empty;
                foreach (SLTable t in slws.Tables)
                {
                    if (t.IsNewTable)
                    {
                        if (t.RelationshipID.Length > 0)
                        {
                            // is a modified existing table
                            tdp = (TableDefinitionPart)wsp.GetPartById(t.RelationshipID);
                            tdp.Table = t.ToTable();
                            sRelID = t.RelationshipID;
                        }
                        else
                        {
                            // is a completely new table
                            tdp = wsp.AddNewPart<TableDefinitionPart>();
                            tdp.Table = t.ToTable();
                            sRelID = wsp.GetIdOfPart(tdp);
                        }

                    }
                    else
                    {
                        // if it's an existing table with no modifications,
                        // don't need to do anything to the XML content.
                        tdp = (TableDefinitionPart)wsp.GetPartById(t.RelationshipID);
                        sRelID = t.RelationshipID;
                    }

                    tps.Append(new TablePart() { Id = sRelID });
                }

                // NOTE: SmartTags is deprecated in Open XML SDK 2.5, so have to remove
                // from check below?

                bFound = false;
                oxe = wsp.Worksheet.FirstChild;
                foreach (var child in wsp.Worksheet.ChildElements)
                {
                    // start with SheetData because it's a required child element
                    if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                        || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                        || child is SortState || child is DataConsolidate || child is CustomSheetViews
                        || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                        || child is DataValidations || child is Hyperlinks || child is PrintOptions
                        || child is PageMargins || child is PageSetup || child is HeaderFooter
                        || child is RowBreaks || child is ColumnBreaks || child is CustomProperties
                        || child is CellWatches || child is IgnoredErrors
                        //|| child is CellWatches || child is IgnoredErrors || child is SmartTags
                        || child is DocumentFormat.OpenXml.Spreadsheet.Drawing
                        || child is LegacyDrawing || child is LegacyDrawingHeaderFooter
                        || child is Picture || child is OleObjects || child is Controls
                        || child is WebPublishItems)
                    {
                        oxe = child;
                        bFound = true;
                    }
                }

                if (bFound)
                {
                    wsp.Worksheet.InsertAfter(tps, oxe);
                }
                else
                {
                    wsp.Worksheet.PrependChild(tps);
                }
            }
            #endregion

            #region 2010 Conditional formatting, Sparklines and possibly other extensions
            WorksheetExtensionList wsextlist;
            WorksheetExtension wsext;
            List<WorksheetExtension> listExtensions = new List<WorksheetExtension>();
            slws.RefreshSparklineGroups();

            if (wsp.Worksheet.Elements<WorksheetExtensionList>().Count() > 0)
            {
                wsextlist = wsp.Worksheet.Elements<WorksheetExtensionList>().First();
                foreach (var wsextchild in wsextlist.ChildElements)
                {
                    if (wsextchild is WorksheetExtension)
                    {
                        wsext = (WorksheetExtension)wsextchild;
                        wsext.RemoveAllChildren<X14.ConditionalFormattings>();
                        wsext.RemoveAllChildren<X14.SparklineGroups>();
                        // there might be other extension types, like slicers (erhmahgerd...).
                        if (wsext.ChildElements.Count > 0)
                        {
                            listExtensions.Add((WorksheetExtension)wsext.CloneNode(true));
                        }
                    }
                }
                wsp.Worksheet.RemoveAllChildren<WorksheetExtensionList>();
            }

            if (slws.ConditionalFormattings2010.Count > 0
                || slws.SparklineGroups.Count > 0
                || listExtensions.Count > 0)
            {
                wsextlist = new WorksheetExtensionList();
                foreach (WorksheetExtension ext in listExtensions)
                {
                    // be extra safe by cloning again to avoid pass-by-reference. Deeply.
                    wsextlist.Append((WorksheetExtension)ext.CloneNode(true));
                }

                if (slws.ConditionalFormattings2010.Count > 0)
                {
                    // this is important! Apparently extensions are tied to a URI that Microsoft uses.
                    wsext = new WorksheetExtension() { Uri = SLConstants.ConditionalFormattingExtensionUri };
                    wsext.AddNamespaceDeclaration("x14", SLConstants.NamespaceX14);
                    X14.ConditionalFormattings cfs = new X14.ConditionalFormattings();
                    foreach (SLConditionalFormatting2010 cfr2010 in slws.ConditionalFormattings2010)
                    {
                        cfs.Append(cfr2010.ToConditionalFormatting());
                    }
                    wsext.Append(cfs);
                    wsextlist.Append(wsext);
                }

                if (slws.SparklineGroups.Count > 0)
                {
                    // this is important! Apparently extensions are tied to a URI that Microsoft uses.
                    wsext = new WorksheetExtension() { Uri = SLConstants.SparklineExtensionUri };
                    wsext.AddNamespaceDeclaration("x14", SLConstants.NamespaceX14);
                    X14.SparklineGroups spkgrps = new X14.SparklineGroups();
                    spkgrps.AddNamespaceDeclaration("xm", SLConstants.NamespaceXm);
                    foreach (SLSparklineGroup spkgrp in slws.SparklineGroups)
                    {
                        spkgrps.Append(spkgrp.ToSparklineGroup());
                    }
                    wsext.Append(spkgrps);
                    wsextlist.Append(wsext);
                }

                // WorksheetExtensionList is the very last element possible.
                // So we can just append it because everything else is in front.
                wsp.Worksheet.Append(wsextlist);
            }
            #endregion

            wsp.Worksheet.Save();
            // end of writing for existing worksheet
        }
        else
        {
            // start of writing for new worksheet
            WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();
            gsSelectedWorksheetRelationshipID = wbp.GetIdOfPart(wsp);
            foreach (SLSheet s in slwb.Sheets)
            {
                if (s.Name.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    s.Id = gsSelectedWorksheetRelationshipID;
                    break;
                }
            }

            if (slws.ForceCustomRowColumnDimensionsSplitting)
            {
                slws.ToggleCustomRowColumnDimension(true);
            }

            List<OpenXmlAttribute> oxa;
            OpenXmlWriter oxw = OpenXmlWriter.Create(wsp);

            oxa = new List<OpenXmlAttribute>();
            oxa.Add(new OpenXmlAttribute("xmlns:r", null, SLConstants.NamespaceRelationships));

            if (slws.ConditionalFormattings2010.Count > 0
                || slws.SparklineGroups.Count > 0)
            {
                oxa.Add(new OpenXmlAttribute("xmlns:x14", null, SLConstants.NamespaceX14));
                oxa.Add(new OpenXmlAttribute("xmlns:xm", null, SLConstants.NamespaceXm));
                oxa.Add(new OpenXmlAttribute("xmlns:mc", null, SLConstants.NamespaceMc));
                oxa.Add(new OpenXmlAttribute("xmlns:x14ac", null, SLConstants.NamespaceX14ac));
                oxa.Add(new OpenXmlAttribute("mc", "Ignorable", SLConstants.NamespaceMc, "x14ac"));
                oxw.WriteStartElement(new Worksheet(), oxa);
            }
            else
            {
                oxw.WriteStartElement(new Worksheet(), oxa);
            }

            if (slws.PageSettings.HasSheetProperties)
            {
                oxw.WriteElement(slws.PageSettings.SheetProperties.ToSheetProperties());
            }

            oxw.WriteElement(new SheetDimension() { Reference = sDimensionCellRange });

            if (slws.SheetViews.Count > 0)
            {
                oxw.WriteStartElement(new SheetViews());

                foreach (SLSheetView sv in slws.SheetViews)
                {
                    oxw.WriteElement(sv.ToSheetView());
                }

                oxw.WriteEndElement();
            }

            oxa = new List<OpenXmlAttribute>();
            if (slws.SheetFormatProperties.BaseColumnWidth != null && slws.SheetFormatProperties.BaseColumnWidth.Value != 8)
            {
                oxa.Add(new OpenXmlAttribute("baseColWidth", null, slws.SheetFormatProperties.BaseColumnWidth.Value.ToString(CultureInfo.InvariantCulture)));
            }
            if (slws.SheetFormatProperties.HasDefaultColumnWidth)
            {
                oxa.Add(new OpenXmlAttribute("defaultColWidth", null, slws.SheetFormatProperties.DefaultColumnWidth.ToString(CultureInfo.InvariantCulture)));
            }
            oxa.Add(new OpenXmlAttribute("defaultRowHeight", null, slws.SheetFormatProperties.DefaultRowHeight.ToString("0.####", CultureInfo.InvariantCulture)));
            if (slws.SheetFormatProperties.CustomHeight != null && slws.SheetFormatProperties.CustomHeight.Value)
            {
                oxa.Add(new OpenXmlAttribute("customHeight", null, "1"));
            }
            if (slws.SheetFormatProperties.ZeroHeight != null && slws.SheetFormatProperties.ZeroHeight.Value)
            {
                oxa.Add(new OpenXmlAttribute("zeroHeight", null, "1"));
            }
            if (slws.SheetFormatProperties.ThickTop != null && slws.SheetFormatProperties.ThickTop.Value)
            {
                oxa.Add(new OpenXmlAttribute("thickTop", null, "1"));
            }
            if (slws.SheetFormatProperties.ThickBottom != null && slws.SheetFormatProperties.ThickBottom.Value)
            {
                oxa.Add(new OpenXmlAttribute("thickBottom", null, "1"));
            }
            if (slws.SheetFormatProperties.OutlineLevelRow != null && slws.SheetFormatProperties.OutlineLevelRow.Value > 0)
            {
                oxa.Add(new OpenXmlAttribute("outlineLevelRow", null, slws.SheetFormatProperties.OutlineLevelRow.Value.ToString(CultureInfo.InvariantCulture)));
            }
            if (slws.SheetFormatProperties.OutlineLevelColumn != null && slws.SheetFormatProperties.OutlineLevelColumn.Value > 0)
            {
                oxa.Add(new OpenXmlAttribute("outlineLevelCol", null, slws.SheetFormatProperties.OutlineLevelColumn.Value.ToString(CultureInfo.InvariantCulture)));
            }
            oxw.WriteStartElement(new SheetFormatProperties(), oxa);
            oxw.WriteEndElement();

            #region Filling Columns
            if (slws.ColumnProperties.Count > 0)
            {
                oxw.WriteStartElement(new Columns());

                int iPreviousColumnIndex = listColumnIndex[0];
                int iCurrentColumnIndex = iPreviousColumnIndex;
                string sCollectiveColumnData = string.Empty;
                string sCurrentColumnData = string.Empty;
                int colmin, colmax;
                colmin = colmax = iCurrentColumnIndex;
                cp = slws.ColumnProperties[iCurrentColumnIndex];
                sCollectiveColumnData = cp.ToHash();

                oxa = new List<OpenXmlAttribute>();
                oxa.Add(new OpenXmlAttribute("min", null, colmin.ToString(CultureInfo.InvariantCulture)));
                // max is left to the end because we're calculating it
                //oxa.Add(new OpenXmlAttribute("max", null, colmax.ToString(CultureInfo.InvariantCulture)));
                if (cp.HasWidth)
                {
                    oxa.Add(new OpenXmlAttribute("width", null, cp.Width.ToString(CultureInfo.InvariantCulture)));
                    oxa.Add(new OpenXmlAttribute("customWidth", null, "1"));
                }
                else
                {
                    oxa.Add(new OpenXmlAttribute("width", null, slws.SheetFormatProperties.DefaultColumnWidth.ToString(CultureInfo.InvariantCulture)));
                }
                if (cp.StyleIndex > 0) oxa.Add(new OpenXmlAttribute("style", null, cp.StyleIndex.ToString(CultureInfo.InvariantCulture)));
                if (cp.Hidden != false) oxa.Add(new OpenXmlAttribute("hidden", null, "1"));
                if (cp.BestFit != false) oxa.Add(new OpenXmlAttribute("bestFit", null, "1"));
                if (cp.Phonetic != false) oxa.Add(new OpenXmlAttribute("phonetic", null, "1"));
                if (cp.OutlineLevel > 0) oxa.Add(new OpenXmlAttribute("outlineLevel", null, cp.OutlineLevel.ToString(CultureInfo.InvariantCulture)));
                if (cp.Collapsed != false) oxa.Add(new OpenXmlAttribute("collapsed", null, "1"));

                for (i = 1; i < listColumnIndex.Count; ++i)
                {
                    iPreviousColumnIndex = iCurrentColumnIndex;
                    iCurrentColumnIndex = listColumnIndex[i];
                    cp = slws.ColumnProperties[iCurrentColumnIndex];
                    sCurrentColumnData = cp.ToHash();

                    if ((iCurrentColumnIndex != (iPreviousColumnIndex + 1)) || (sCollectiveColumnData != sCurrentColumnData))
                    {
                        oxa.Add(new OpenXmlAttribute("max", null, colmax.ToString(CultureInfo.InvariantCulture)));
                        oxw.WriteStartElement(new Column(), oxa);
                        oxw.WriteEndElement();

                        colmin = iCurrentColumnIndex;
                        colmax = iCurrentColumnIndex;
                        sCollectiveColumnData = sCurrentColumnData;

                        oxa = new List<OpenXmlAttribute>();
                        oxa.Add(new OpenXmlAttribute("min", null, colmin.ToString(CultureInfo.InvariantCulture)));
                        if (cp.HasWidth)
                        {
                            oxa.Add(new OpenXmlAttribute("width", null, cp.Width.ToString(CultureInfo.InvariantCulture)));
                            oxa.Add(new OpenXmlAttribute("customWidth", null, "1"));
                        }
                        else
                        {
                            oxa.Add(new OpenXmlAttribute("width", null, slws.SheetFormatProperties.DefaultColumnWidth.ToString(CultureInfo.InvariantCulture)));
                        }
                        if (cp.StyleIndex > 0) oxa.Add(new OpenXmlAttribute("style", null, cp.StyleIndex.ToString(CultureInfo.InvariantCulture)));
                        if (cp.Hidden != false) oxa.Add(new OpenXmlAttribute("hidden", null, "1"));
                        if (cp.BestFit != false) oxa.Add(new OpenXmlAttribute("bestFit", null, "1"));
                        if (cp.Phonetic != false) oxa.Add(new OpenXmlAttribute("phonetic", null, "1"));
                        if (cp.OutlineLevel > 0) oxa.Add(new OpenXmlAttribute("outlineLevel", null, cp.OutlineLevel.ToString(CultureInfo.InvariantCulture)));
                        if (cp.Collapsed != false) oxa.Add(new OpenXmlAttribute("collapsed", null, "1"));
                    }
                    else
                    {
                        colmax = iCurrentColumnIndex;
                    }
                }

                // there's always a "leftover" column
                oxa.Add(new OpenXmlAttribute("max", null, colmax.ToString(CultureInfo.InvariantCulture)));
                oxw.WriteStartElement(new Column(), oxa);
                oxw.WriteEndElement();

                oxw.WriteEndElement();
            }
            #endregion

            oxw.WriteStartElement(new SheetData());

            foreach (int rowkey in listRowKeys)
            {
                oxa = new List<OpenXmlAttribute>();
                oxa.Add(new OpenXmlAttribute("r", null, rowkey.ToString(CultureInfo.InvariantCulture)));
                if (slws.RowProperties.ContainsKey(rowkey))
                {
                    rp = slws.RowProperties[rowkey];
                    if (rp.StyleIndex > 0)
                    {
                        oxa.Add(new OpenXmlAttribute("s", null, rp.StyleIndex.ToString(CultureInfo.InvariantCulture)));
                        oxa.Add(new OpenXmlAttribute("customFormat", null, "1"));
                    }
                    if (rp.HasHeight)
                    {
                        oxa.Add(new OpenXmlAttribute("ht", null, rp.Height.ToString(CultureInfo.InvariantCulture)));
                    }
                    if (rp.Hidden != false)
                    {
                        oxa.Add(new OpenXmlAttribute("hidden", null, "1"));
                    }
                    if (rp.CustomHeight)
                    {
                        oxa.Add(new OpenXmlAttribute("customHeight", null, "1"));
                    }
                    if (rp.OutlineLevel > 0)
                    {
                        oxa.Add(new OpenXmlAttribute("outlineLevel", null, rp.OutlineLevel.ToString(CultureInfo.InvariantCulture)));
                    }
                    if (rp.Collapsed != false)
                    {
                        oxa.Add(new OpenXmlAttribute("collapsed", null, "1"));
                    }
                    if (rp.ThickTop != false)
                    {
                        oxa.Add(new OpenXmlAttribute("thickTop", null, "1"));
                    }
                    if (rp.ThickBottom != false)
                    {
                        oxa.Add(new OpenXmlAttribute("thickBot", null, "1"));
                    }
                    if (rp.ShowPhonetic != false)
                    {
                        oxa.Add(new OpenXmlAttribute("ph", null, "1"));
                    }
                }
                oxw.WriteStartElement(new Row(), oxa);

                if (slws.CellWarehouse.Cells.ContainsKey(rowkey))
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    listColumnKeys.Sort();
                    foreach (int colkey in listColumnKeys)
                    {
                        c = slws.CellWarehouse.Cells[rowkey][colkey];

                        oxa = new List<OpenXmlAttribute>();
                        oxa.Add(new OpenXmlAttribute("r", null, SLTool.ToCellReference(rowkey, colkey)));
                        if (c.StyleIndex > 0)
                        {
                            oxa.Add(new OpenXmlAttribute("s", null, c.StyleIndex.ToString(CultureInfo.InvariantCulture)));
                        }

                        // number type is default
                        switch (c.DataType)
                        {
                            case CellValues.Boolean:
                                oxa.Add(new OpenXmlAttribute("t", null, "b"));
                                break;
                            case CellValues.Date:
                                oxa.Add(new OpenXmlAttribute("t", null, "d"));
                                break;
                            case CellValues.Error:
                                oxa.Add(new OpenXmlAttribute("t", null, "e"));
                                break;
                            case CellValues.InlineString:
                                oxa.Add(new OpenXmlAttribute("t", null, "inlineStr"));
                                break;
                            case CellValues.SharedString:
                                oxa.Add(new OpenXmlAttribute("t", null, "s"));
                                break;
                            case CellValues.String:
                                oxa.Add(new OpenXmlAttribute("t", null, "str"));
                                break;
                        }

                        if (c.CellMetaIndex > 0)
                        {
                            oxa.Add(new OpenXmlAttribute("cm", null, c.CellMetaIndex.ToString(CultureInfo.InvariantCulture)));
                        }
                        if (c.ValueMetaIndex > 0)
                        {
                            oxa.Add(new OpenXmlAttribute("vm", null, c.ValueMetaIndex.ToString(CultureInfo.InvariantCulture)));
                        }
                        if (c.ShowPhonetic != false)
                        {
                            oxa.Add(new OpenXmlAttribute("ph", null, "1"));
                        }
                        oxw.WriteStartElement(new Cell(), oxa);
                        if (c.CellFormula != null)
                        {
                            oxw.WriteElement(c.CellFormula.ToCellFormula());
                        }

                        if (c.CellText != null)
                        {
                            if (c.CellText.Length > 0)
                            {
                                if (c.ToPreserveSpace)
                                {
                                    oxw.WriteElement(new CellValue(c.CellText)
                                    {
                                        Space = SpaceProcessingModeValues.Preserve
                                    });
                                }
                                else
                                {
                                    oxw.WriteElement(new CellValue(c.CellText));
                                }
                            }
                        }
                        else
                        {
                            if (c.DataType == CellValues.Number)
                            {
                                oxw.WriteElement(new CellValue(c.NumericValue.ToString(CultureInfo.InvariantCulture)));
                            }
                            else if (c.DataType == CellValues.SharedString)
                            {
                                oxw.WriteElement(new CellValue(c.NumericValue.ToString("f0", CultureInfo.InvariantCulture)));
                            }
                            else if (c.DataType == CellValues.Boolean)
                            {
                                if (c.NumericValue > 0.5) oxw.WriteElement(new CellValue("1"));
                                else oxw.WriteElement(new CellValue("0"));
                            }
                        }
                        oxw.WriteEndElement();
                    }
                }

                // end of Row
                oxw.WriteEndElement();
            }

            // end of SheetData
            oxw.WriteEndElement();

            #region Sheet protection
            if (slws.HasSheetProtection)
            {
                oxw.WriteElement(slws.SheetProtection.ToSheetProtection());
            }
            #endregion

            #region AutoFilter
            if (slws.HasAutoFilter)
            {
                oxw.WriteElement(slws.AutoFilter.ToAutoFilter());
            }
            #endregion

            #region Filling merge cells
            if (slws.MergeCells.Count > 0)
            {
                oxw.WriteStartElement(new MergeCells() { Count = (uint)slws.MergeCells.Count });
                for (i = 0; i < slws.MergeCells.Count; ++i)
                {
                    oxw.WriteElement(slws.MergeCells[i].ToMergeCell());
                }
                oxw.WriteEndElement();
            }
            #endregion

            #region Conditional Formatting
            if (slws.ConditionalFormattings.Count > 0)
            {
                for (i = 0; i < slws.ConditionalFormattings.Count; ++i)
                {
                    oxw.WriteElement(slws.ConditionalFormattings[i].ToConditionalFormatting());
                }
            }
            #endregion

            #region DataValidations
            if (slws.DataValidations.Count > 0)
            {
                DataValidations dvs = new DataValidations();
                if (slws.DataValidationDisablePrompts) dvs.DisablePrompts = slws.DataValidationDisablePrompts;
                if (slws.DataValidationXWindow != null) dvs.XWindow = slws.DataValidationXWindow.Value;
                if (slws.DataValidationYWindow != null) dvs.YWindow = slws.DataValidationYWindow.Value;
                dvs.Count = (uint)slws.DataValidations.Count;

                foreach (SLDataValidation dv in slws.DataValidations)
                {
                    dvs.Append(dv.ToDataValidation());
                }

                oxw.WriteElement(dvs);
            }
            #endregion

            #region Hyperlinks
            if (slws.Hyperlinks.Count > 0)
            {
                Hyperlinks hls = new Hyperlinks();
                HyperlinkRelationship hlrel;
                foreach (SLHyperlink hl in slws.Hyperlinks)
                {
                    if (hl.IsExternal && hl.IsNew)
                    {
                        hlrel = wsp.AddHyperlinkRelationship(new Uri(hl.HyperlinkUri, hl.HyperlinkUriKind), true);
                        hl.Id = hlrel.Id;
                    }
                    hls.Append(hl.ToHyperlink());
                }

                oxw.WriteElement(hls);
            }
            #endregion

            #region PrintOptions
            if (slws.PageSettings.HasPrintOptions)
            {
                oxw.WriteElement(slws.PageSettings.ExportPrintOptions());
            }
            #endregion

            #region PageMargins
            if (slws.PageSettings.HasPageMargins)
            {
                oxw.WriteElement(slws.PageSettings.ExportPageMargins());
            }
            #endregion

            #region PageSetup
            if (slws.PageSettings.HasPageSetup)
            {
                oxw.WriteElement(slws.PageSettings.ExportPageSetup());
            }
            #endregion

            #region HeaderFooter
            if (slws.PageSettings.HasHeaderFooter)
            {
                oxw.WriteElement(slws.PageSettings.ExportHeaderFooter());
            }
            #endregion

            #region RowBreaks
            if (slws.RowBreaks.Count > 0)
            {
                List<int> bkkeys = slws.RowBreaks.Keys.ToList<int>();
                bkkeys.Sort();

                // if it's a new worksheet, then all breaks are considered manual
                oxw.WriteStartElement(new RowBreaks()
                {
                    Count = (uint)slws.RowBreaks.Count,
                    ManualBreakCount = (uint)slws.RowBreaks.Count
                });
                foreach (int bkindex in bkkeys)
                {
                    oxw.WriteElement(slws.RowBreaks[bkindex].ToBreak());
                }
                oxw.WriteEndElement();
            }
            #endregion

            #region ColumnBreaks
            if (slws.ColumnBreaks.Count > 0)
            {
                List<int> bkkeys = slws.ColumnBreaks.Keys.ToList<int>();
                bkkeys.Sort();

                // if it's a new worksheet, then all breaks are considered manual
                oxw.WriteStartElement(new ColumnBreaks()
                {
                    Count = (uint)slws.ColumnBreaks.Count,
                    ManualBreakCount = (uint)slws.ColumnBreaks.Count
                });
                foreach (int bkindex in bkkeys)
                {
                    oxw.WriteElement(slws.ColumnBreaks[bkindex].ToBreak());
                }
                oxw.WriteEndElement();
            }
            #endregion

            #region Drawing
            // these are "new" charts and pictures added
            if (slws.Charts.Count > 0 || slws.Pictures.Count > 0)
            {
                DrawingsPart dp = wsp.AddNewPart<DrawingsPart>();
                dp.WorksheetDrawing = new Xdr.WorksheetDrawing();
                dp.WorksheetDrawing.AddNamespaceDeclaration("xdr", SLConstants.NamespaceXdr);
                dp.WorksheetDrawing.AddNamespaceDeclaration("a", SLConstants.NamespaceA);

                oxw.WriteElement(new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = wsp.GetIdOfPart(dp) });

                WriteImageParts(dp);
            }
            #endregion

            #region LegacyDrawing
            // these are "new" comments added
            if (slws.Comments.Count > 0)
            {
                WorksheetCommentsPart wcp = wsp.AddNewPart<WorksheetCommentsPart>();
                VmlDrawingPart vdp = wsp.AddNewPart<VmlDrawingPart>();
                WriteCommentPart(wcp, vdp);

                oxw.WriteElement(new LegacyDrawing() { Id = wsp.GetIdOfPart(vdp) });
            }
            #endregion

            #region Picture
            if (slws.BackgroundPictureDataIsInFile != null)
            {
                ImagePart imgp = wsp.AddImagePart(slws.BackgroundPictureImagePartType);
                if (slws.BackgroundPictureDataIsInFile.Value)
                {
                    using (FileStream fs = new FileStream(slws.BackgroundPictureFileName, FileMode.Open))
                    {
                        imgp.FeedData(fs);
                    }
                }
                else
                {
                    using (MemoryStream ms = new MemoryStream(slws.BackgroundPictureByteData))
                    {
                        imgp.FeedData(ms);
                    }
                }

                oxw.WriteElement(new Picture() { Id = wsp.GetIdOfPart(imgp) });
            }
            #endregion

            if (slws.Tables.Count > 0)
            {
                // If it's a new worksheet, ALL tables are new tables...
                oxw.WriteStartElement(new TableParts() { Count = (uint)slws.Tables.Count });
                TableDefinitionPart tdp;
                foreach (SLTable t in slws.Tables)
                {
                    tdp = wsp.AddNewPart<TableDefinitionPart>();
                    tdp.Table = t.ToTable();
                    oxw.WriteElement(new TablePart() { Id = wsp.GetIdOfPart(tdp) });
                }
                oxw.WriteEndElement();
            }

            #region 2010 Conditional formatting, Sparklines and possibly other extensions
            slws.RefreshSparklineGroups();

            if (slws.ConditionalFormattings2010.Count > 0
                || slws.SparklineGroups.Count > 0)
            {
                oxw.WriteStartElement(new WorksheetExtensionList());

                if (slws.ConditionalFormattings2010.Count > 0)
                {
                    oxa = new List<OpenXmlAttribute>();
                    oxa.Add(new OpenXmlAttribute("xmlns:x14", null, SLConstants.NamespaceX14));
                    // this is important! Apparently extensions are tied to a URI that Microsoft uses.
                    oxa.Add(new OpenXmlAttribute("uri", null, SLConstants.ConditionalFormattingExtensionUri));
                    oxw.WriteStartElement(new WorksheetExtension(), oxa);

                    oxw.WriteStartElement(new X14.ConditionalFormattings());
                    foreach (SLConditionalFormatting2010 cf2010 in slws.ConditionalFormattings2010)
                    {
                        oxw.WriteElement(cf2010.ToConditionalFormatting());
                    }
                    oxw.WriteEndElement();

                    oxw.WriteEndElement();
                }

                if (slws.SparklineGroups.Count > 0)
                {
                    oxa = new List<OpenXmlAttribute>();
                    oxa.Add(new OpenXmlAttribute("xmlns:x14", null, SLConstants.NamespaceX14));
                    // this is important! Apparently extensions are tied to a URI that Microsoft uses.
                    oxa.Add(new OpenXmlAttribute("uri", null, SLConstants.SparklineExtensionUri));
                    oxw.WriteStartElement(new WorksheetExtension(), oxa);

                    oxa = new List<OpenXmlAttribute>();
                    oxa.Add(new OpenXmlAttribute("xmlns:xm", null, SLConstants.NamespaceXm));
                    oxw.WriteStartElement(new X14.SparklineGroups(), oxa);
                    foreach (SLSparklineGroup spkgrp in slws.SparklineGroups)
                    {
                        oxw.WriteElement(spkgrp.ToSparklineGroup());
                    }
                    oxw.WriteEndElement();

                    oxw.WriteEndElement();
                }

                oxw.WriteEndElement();
            }
            #endregion

            oxw.WriteEndElement();
            oxw.Dispose();
            // end of writing for new worksheet
        }
    }

    private void WriteWorkbook()
    {
        int i = 0;

        if (!IsNewSpreadsheet)
        {
            wbp.Workbook.FileVersion = new FileVersion() { ApplicationName = SLConstants.ApplicationName };

            if (slwb.WorkbookProperties.HasWorkbookProperties)
            {
                wbp.Workbook.WorkbookProperties = slwb.WorkbookProperties.ToWorkbookProperties();
            }
            else
            {
                wbp.Workbook.WorkbookProperties = null;
            }

            if (slwb.WorkbookViews.Count > 0)
            {
                wbp.Workbook.BookViews = new BookViews();
                for (i = 0; i < slwb.WorkbookViews.Count; ++i)
                {
                    wbp.Workbook.BookViews.Append(slwb.WorkbookViews[i].ToWorkbookView());
                }
            }
            else
            {
                wbp.Workbook.BookViews = null;
            }

            Sheets sheets = new Sheets();
            for (i = 0; i < slwb.Sheets.Count; ++i)
            {
                sheets.Append(slwb.Sheets[i].ToSheet());
            }
            wbp.Workbook.Sheets = sheets;

            if (slwb.DefinedNames.Count > 0)
            {
                DefinedNames dns = new DefinedNames();
                for (i = 0; i < slwb.DefinedNames.Count; ++i)
                {
                    dns.Append(slwb.DefinedNames[i].ToDefinedName());
                }
                wbp.Workbook.DefinedNames = dns;
            }
            else
            {
                // just assign a new DefinedNames() to clear out the existing one
                if (wbp.Workbook.DefinedNames != null) wbp.Workbook.DefinedNames = new DefinedNames();
            }

            if (slwb.CalculationCells.Count > 0)
            {
                if (wbp.CalculationChainPart == null) wbp.AddNewPart<CalculationChainPart>();
                wbp.CalculationChainPart.CalculationChain = new CalculationChain();
                CalculationCell cc;
                int iCurrentSheetId = 0;
                for (i = 0; i < slwb.CalculationCells.Count; ++i)
                {
                    cc = slwb.CalculationCells[i].ToCalculationCell();
                    if (cc.SheetId.Value == iCurrentSheetId) cc.SheetId = null;
                    else iCurrentSheetId = cc.SheetId.Value;

                    wbp.CalculationChainPart.CalculationChain.Append(cc);
                }
                wbp.CalculationChainPart.CalculationChain.Save();

                if (wbp.Workbook.CalculationProperties == null)
                {
                    wbp.Workbook.CalculationProperties = new CalculationProperties()
                    {
                        CalculationId = SLConstants.CalculationId
                    };
                }
            }
            else
            {
                if (wbp.CalculationChainPart != null)
                {
                    wbp.DeletePart(wbp.GetIdOfPart(wbp.CalculationChainPart));
                }
            }

            wbp.Workbook.Save();
        }
        else
        {
            using (OpenXmlWriter oxw = OpenXmlWriter.Create(wbp))
            {
                List<OpenXmlAttribute> oxa = new List<OpenXmlAttribute>();
                oxa.Add(new OpenXmlAttribute("xmlns:r", null, SLConstants.NamespaceRelationships));
                oxw.WriteStartElement(new Workbook(), oxa);

                oxw.WriteElement(new FileVersion() { ApplicationName = SLConstants.ApplicationName });

                if (slwb.WorkbookProperties.HasWorkbookProperties)
                {
                    oxw.WriteElement(slwb.WorkbookProperties.ToWorkbookProperties());
                }

                if (slwb.WorkbookViews.Count > 0)
                {
                    oxw.WriteStartElement(new BookViews());
                    for (i = 0; i < slwb.WorkbookViews.Count; ++i)
                    {
                        oxw.WriteElement(slwb.WorkbookViews[i].ToWorkbookView());
                    }
                    oxw.WriteEndElement();
                }

                SLSheet sheet;
                oxw.WriteStartElement(new Sheets());
                for (i = 0; i < slwb.Sheets.Count; ++i)
                {
                    sheet = slwb.Sheets[i];
                    if (sheet.State == SheetStateValues.Visible)
                    {
                        oxw.WriteElement(new Sheet()
                        {
                            Name = sheet.Name,
                            SheetId = sheet.SheetId,
                            Id = sheet.Id
                        });
                    }
                    else
                    {
                        oxw.WriteElement(new Sheet()
                        {
                            Name = sheet.Name,
                            SheetId = sheet.SheetId,
                            State = sheet.State,
                            Id = sheet.Id
                        });
                    }
                }
                oxw.WriteEndElement();

                if (slwb.DefinedNames.Count > 0)
                {
                    oxw.WriteStartElement(new DefinedNames());
                    for (i = 0; i < slwb.DefinedNames.Count; ++i)
                    {
                        oxw.WriteElement(slwb.DefinedNames[i].ToDefinedName());
                    }
                    oxw.WriteEndElement();
                }

                if (slwb.CalculationCells.Count > 0)
                {
                    oxw.WriteElement(new CalculationProperties() { CalculationId = SLConstants.CalculationId });
                }

                // workbook
                oxw.WriteEndElement();
            }

            if (slwb.CalculationCells.Count > 0)
            {
                wbp.AddNewPart<CalculationChainPart>();
                using (OpenXmlWriter oxw = OpenXmlWriter.Create(wbp.CalculationChainPart))
                {
                    oxw.WriteStartElement(new CalculationChain());
                    for (i = 0; i < slwb.CalculationCells.Count; ++i)
                    {
                        oxw.WriteElement(slwb.CalculationCells[i].ToCalculationCell());
                    }
                    oxw.WriteEndElement();
                }
            }
        }
    }

    private void LoadDocumentProperties()
    {
        if (xl.CoreFilePropertiesPart != null)
        {
            Stream stream = xl.CoreFilePropertiesPart.GetStream();
            XDocument xdoc = XDocument.Load(XmlReader.Create(stream));

            foreach (XElement xelem in xdoc.Descendants())
            {
                switch (xelem.Name.LocalName)
                {
                    case "category":
                        this.DocumentProperties.Category = xelem.Value;
                        break;
                    case "contentStatus":
                        this.DocumentProperties.ContentStatus = xelem.Value;
                        break;
                    case "created":
                        this.DocumentProperties.Created = xelem.Value;
                        break;
                    case "creator":
                        this.DocumentProperties.Creator = xelem.Value;
                        break;
                    case "description":
                        this.DocumentProperties.Description = xelem.Value;
                        break;
                    case "identifier":
                        this.DocumentProperties.Identifier = xelem.Value;
                        break;
                    case "keywords":
                        this.DocumentProperties.Keywords = xelem.Value;
                        break;
                    case "language":
                        this.DocumentProperties.Language = xelem.Value;
                        break;
                    case "lastModifiedBy":
                        this.DocumentProperties.LastModifiedBy = xelem.Value;
                        break;
                    case "lastPrinted":
                        this.DocumentProperties.LastPrinted = xelem.Value;
                        break;
                    case "modified":
                        this.DocumentProperties.Modified = xelem.Value;
                        break;
                    case "revision":
                        this.DocumentProperties.Revision = xelem.Value;
                        break;
                    case "subject":
                        this.DocumentProperties.Subject = xelem.Value;
                        break;
                    case "title":
                        this.DocumentProperties.Title = xelem.Value;
                        break;
                    case "version":
                        this.DocumentProperties.Version = xelem.Value;
                        break;
                }
            }

            stream.Close();
        }
    }

    private void WriteDocumentProperties()
    {
        if (xl.CoreFilePropertiesPart == null)
        {
            xl.AddCoreFilePropertiesPart();
        }

        using (XmlWriter xw = XmlWriter.Create(xl.CoreFilePropertiesPart.GetStream(FileMode.Create, FileAccess.Write)))
        {
            xw.WriteStartDocument(true);
            xw.WriteStartElement("cp", "coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            xw.WriteAttributeString("xmlns", "dc", null, "http://purl.org/dc/elements/1.1/");
            xw.WriteAttributeString("xmlns", "dcmitype", null, "http://purl.org/dc/dcmitype/");
            xw.WriteAttributeString("xmlns", "dcterms", null, "http://purl.org/dc/terms/");
            xw.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");

            if (this.DocumentProperties.Category.Length > 0)
            {
                xw.WriteElementString("cp", "category", null, this.DocumentProperties.Category);
            }

            if (this.DocumentProperties.ContentStatus.Length > 0)
            {
                xw.WriteElementString("cp", "contentStatus", null, this.DocumentProperties.ContentStatus);
            }

            if (this.DocumentProperties.Created.Length > 0)
            {
                xw.WriteStartElement("dcterms", "created", null);
                xw.WriteAttributeString("xsi", "type", null, "dcterms:W3CDTF");
                xw.WriteString(this.DocumentProperties.Created);
                xw.WriteEndElement();
            }

            if (this.DocumentProperties.Creator.Length > 0)
            {
                xw.WriteElementString("dc", "creator", null, this.DocumentProperties.Creator);
            }

            if (this.DocumentProperties.Description.Length > 0)
            {
                xw.WriteElementString("dc", "description", null, this.DocumentProperties.Description);
            }

            if (this.DocumentProperties.Identifier.Length > 0)
            {
                xw.WriteElementString("dc", "identifier", null, this.DocumentProperties.Identifier);
            }

            if (this.DocumentProperties.Keywords.Length > 0)
            {
                xw.WriteElementString("cp", "keywords", null, this.DocumentProperties.Keywords);
            }

            if (this.DocumentProperties.Language.Length > 0)
            {
                xw.WriteElementString("dc", "language", null, this.DocumentProperties.Language);
            }

            if (this.DocumentProperties.LastModifiedBy.Length > 0)
            {
                xw.WriteElementString("cp", "lastModifiedBy", null, this.DocumentProperties.LastModifiedBy);
            }

            if (this.DocumentProperties.LastPrinted.Length > 0)
            {
                xw.WriteElementString("cp", "lastPrinted", null, this.DocumentProperties.LastPrinted);
            }

            // we're modifying, so we're ignoring the existing one
            xw.WriteStartElement("dcterms", "modified", null);
            xw.WriteAttributeString("xsi", "type", null, "dcterms:W3CDTF");
            xw.WriteString(DateTime.UtcNow.ToString(SLConstants.W3CDTF));
            xw.WriteEndElement();

            if (this.DocumentProperties.Revision.Length > 0)
            {
                xw.WriteElementString("cp", "revision", null, this.DocumentProperties.Revision);
            }

            if (this.DocumentProperties.Subject.Length > 0)
            {
                xw.WriteElementString("dc", "subject", null, this.DocumentProperties.Subject);
            }

            if (this.DocumentProperties.Title.Length > 0)
            {
                xw.WriteElementString("dc", "title", null, this.DocumentProperties.Title);
            }

            if (this.DocumentProperties.Version.Length > 0)
            {
                xw.WriteElementString("cp", "version", null, this.DocumentProperties.Version);
            }

            xw.WriteEndElement();
            xw.WriteEndDocument();
            xw.Close();
        }
    }

    /// <summary>
    /// Make sure to call this only after all worksheets are done! But before writing the workbook!
    /// </summary>
    private void WriteTabSelectedWorksheet()
    {
        bool bFound = false;
        string sRelID = string.Empty;
        foreach (WorksheetPart wsp in wbp.WorksheetParts)
        {
            #region Worksheets
            sRelID = wbp.GetIdOfPart(wsp);
            if (sRelID.Equals(gsSelectedWorksheetRelationshipID, StringComparison.OrdinalIgnoreCase))
            {
                if (wsp.Worksheet.SheetViews != null)
                {
                    bFound = false;
                    foreach (SheetView sv in wsp.Worksheet.SheetViews)
                    {
                        // need to do for all sheet views?
                        if (sv.WorkbookViewId == 0)
                        {
                            bFound = true;
                            sv.TabSelected = true;
                        }
                    }

                    if (!bFound)
                    {
                        wsp.Worksheet.SheetViews.Append(new SheetView() { TabSelected = true, WorkbookViewId = 0 });
                    }
                }
                else
                {
                    wsp.Worksheet.SheetViews = new SheetViews();
                    wsp.Worksheet.SheetViews.Append(new SheetView() { TabSelected = true, WorkbookViewId = 0 });
                }
            }
            else
            {
                if (wsp.Worksheet.SheetViews != null)
                {
                    foreach (SheetView sv in wsp.Worksheet.SheetViews)
                    {
                        // the default value is false, and setting null generates less XML data
                        sv.TabSelected = null;
                    }
                }
                // no else because we'll just ignore
            }

            // TODO: Custom sheet views?

            wsp.Worksheet.Save();

            #endregion
        }

        #region Chartsheets
        foreach (ChartsheetPart csp in wbp.ChartsheetParts)
        {
            if (csp.Chartsheet.ChartSheetViews != null)
            {
                foreach (ChartSheetView csv in csp.Chartsheet.ChartSheetViews)
                {
                    // the default value is false, and setting null generates less XML data
                    csv.TabSelected = null;
                }
            }
            // no else because we'll just ignore

            // TODO: Custom chart sheet views?

            csp.Chartsheet.Save();
        }
        #endregion

        #region Dialogsheets
        foreach (DialogsheetPart dsp in wbp.DialogsheetParts)
        {
            if (dsp.DialogSheet.SheetViews != null)
            {
                foreach (SheetView sv in dsp.DialogSheet.SheetViews)
                {
                    // the default value is false, and setting null generates less XML data
                    sv.TabSelected = null;
                }
            }
            // no else because we'll just ignore

            // TODO: Custom sheet views?

            dsp.DialogSheet.Save();
        }
        #endregion

        #region Macrosheets
        foreach (MacroSheetPart msp in wbp.MacroSheetParts)
        {
            if (msp.Macrosheet.SheetViews != null)
            {
                foreach (SheetView sv in msp.Macrosheet.SheetViews)
                {
                    // the default value is false, and setting null generates less XML data
                    sv.TabSelected = null;
                }
            }
            // no else because we'll just ignore

            // TODO: Custom sheet views?

            msp.Macrosheet.Save();
        }
        #endregion

        // the currently selected worksheet "must" be visible, right?
        this.ShowWorksheet(gsSelectedWorksheetName);

        // if you think about it, the active tab is always >= the first sheet
        uint? iFirstSheet = null;
        uint? iActiveTab = null;
        int i;

        for (i = 0; i < slwb.Sheets.Count; ++i)
        {
            // there's guaranteed to be at least one visible worksheet
            // because we made it so above.
            if (iFirstSheet == null)
            {
                if (slwb.Sheets[i].State == SheetStateValues.Visible) iFirstSheet = (uint)i;
            }

            // there's guaranteed to be one worksheet with the same name as the 
            // currently selected worksheet.
            if (iActiveTab == null)
            {
                if (slwb.Sheets[i].Name.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    iActiveTab = (uint)i;
                }
            }
        }

        // basically with the double guarantees, the first sheet and active tab cannot be both null.
        // but we check for nullity just in case...
        if (iFirstSheet != null && iActiveTab != null)
        {
            if (slwb.WorkbookViews.Count > 0)
            {
                if (iFirstSheet != null) slwb.WorkbookViews[0].FirstSheet = iFirstSheet.Value;
                if (iActiveTab != null) slwb.WorkbookViews[0].ActiveTab = iActiveTab.Value;
            }
            else
            {
                SLWorkbookView wv = new SLWorkbookView();
                if (iFirstSheet != null) wv.FirstSheet = iFirstSheet.Value;
                if (iActiveTab != null) wv.ActiveTab = iActiveTab.Value;
                slwb.WorkbookViews.Add(wv);
            }
        }

        // I'm not creating a default workbook view because the main info seems to be the position
        // and size of the workbook view within the application (read: Excel).
        // Since I don't know how to calculate it (because I can't *see* the user's screen), I'm
        // going to ignore it. Twips? Good grief...
    }

    private void CloseAndCleanUp()
    {
        WriteSelectedWorksheet();
        WriteTabSelectedWorksheet();
        WriteTheme();
        WriteStylesheet();
        WriteSharedStringTable();

        WriteWorkbook();
        WriteDocumentProperties();
        
        xl.Close();

        // 2020-11-14 .NET Standard update
        // Opening a part more than once for writing results in the error:
        // "Entries cannot be opened multiple times in Update mode"
        // This is a hard limitation from System.IO.Compression
        // Perhaps in future, this soft fix can be set at the point of creating
        // the worksheet parts. But is it still worth it?
        #region LibreOffice and possibly iPhone/iPad fix
        // This will solve LibreOffice not opening documents correctly if document metadata is set.
        // Also, (hopefully) this solves any iPhone/iPad issue with opening documents.
        //string sDocSchema = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        //string sDocFolder = string.Empty;
        //int index;
        //List<string> listUri = new List<string>();
        //using (Package pkg = Package.Open(memstream, FileMode.Open, FileAccess.ReadWrite))
        //{
        //    foreach (PackageRelationship pkgrel in pkg.GetRelationshipsByType(sDocSchema))
        //    {
        //        sDocFolder = pkgrel.TargetUri.OriginalString;
        //        // there should only be one
        //        break;
        //    }

        //    sDocFolder = sDocFolder.TrimStart("/".ToCharArray());
        //    index = sDocFolder.LastIndexOf("/");
        //    if (index <= 0) sDocFolder = string.Empty;
        //    else sDocFolder = sDocFolder.Substring(0, index);

        //    if (sDocFolder.Length > 0)
        //    {
        //        foreach (PackagePart pkgpart in pkg.GetParts())
        //        {
        //            if (pkgpart.Uri.OriginalString.EndsWith(".rels"))
        //            {
        //                listUri.Add(pkgpart.Uri.OriginalString);
        //            }
        //        }

        //        Uri reluri;
        //        string sRelUri;
        //        string sFolder;

        //        int iFolderCount = 0;
        //        int iRelsCase = -1;
        //        string sRelSchema = "http://schemas.openxmlformats.org/package/2006/relationships";
        //        XName targetattr = XName.Get("Target");

        //        foreach (string sUri in listUri)
        //        {
        //            if (sUri.StartsWith("/_rels/"))
        //            {
        //                iRelsCase = 0;
        //            }
        //            else if (sUri.StartsWith(string.Format("/{0}/_rels/", sDocFolder)))
        //            {
        //                iRelsCase = 1;
        //            }
        //            else if (sUri.StartsWith(string.Format("/{0}/", sDocFolder)))
        //            {
        //                iRelsCase = 2;

        //                // Originally "/{docFolder}/drawings/_rels/drawing1.xml.rels"
        //                // Now "drawings/_rels/drawing1.xml.rels"
        //                // +2 for the start and end slashes
        //                sFolder = sUri.Substring(sDocFolder.Length + 2);
        //                index = sFolder.IndexOf("/_rels/");
        //                // We want to get "drawings/", that's why we +1 to include the slash
        //                sFolder = sFolder.Substring(0, index + 1);
        //                // count the number of slashes. A quick and dirty way of counting
        //                // how many folders we are in.
        //                iFolderCount = sFolder.Length - sFolder.Replace("/", "").Length;
        //            }
        //            else
        //            {
        //                // we're not interested in processing anything other than the above
        //                iRelsCase = -1;
        //                continue;
        //            }

        //            reluri = new Uri(sUri, UriKind.Relative);
        //            if (pkg.PartExists(reluri))
        //            {
        //                XDocument xdoc = XDocument.Load(XmlReader.Create(pkg.GetPart(reluri).GetStream()));
        //                foreach (XElement xelem in xdoc.Elements(XName.Get("Relationships", sRelSchema)).Elements(XName.Get("Relationship", sRelSchema)))
        //                {
        //                    sRelUri = xelem.Attribute(targetattr).Value;

        //                    switch (iRelsCase)
        //                    {
        //                        case 0:
        //                            sRelUri = sRelUri.TrimStart("/".ToCharArray());
        //                            break;
        //                        case 1:
        //                            if (sRelUri.StartsWith(string.Format("/{0}/", sDocFolder)))
        //                            {
        //                                // +2 for the start and end slashes
        //                                sRelUri = sRelUri.Substring(sDocFolder.Length + 2);
        //                            }
        //                            break;
        //                        case 2:
        //                            if (sRelUri.StartsWith(string.Format("/{0}/", sDocFolder)))
        //                            {
        //                                // +2 for the start and end slashes
        //                                sRelUri = sRelUri.Substring(sDocFolder.Length + 2);
        //                                for (index = 0; index < iFolderCount; ++index)
        //                                {
        //                                    sRelUri = "../" + sRelUri;
        //                                }
        //                            }
        //                            break;
        //                    }

        //                    xelem.Attribute(targetattr).Value = sRelUri;
        //                }

        //                using (MemoryStream ms = new MemoryStream())
        //                {
        //                    StreamWriter sw = new StreamWriter(ms);
        //                    xdoc.Save(sw);
        //                    pkg.GetPart(reluri).GetStream(FileMode.Create, FileAccess.Write).Write(ms.ToArray(), 0, (int)ms.Length);
        //                    sw.Close();
        //                }
        //            }
        //        }
        //    }
        //    // else the doc folder is empty string, then don't have to do anything.
        //    // Although that seems unlikely...

        //    pkg.Close();
        //}
        #endregion

        this.NullifyInternalDataStores();
    }

    private void NullifyInternalDataStores()
    {
        slwb = null;
        slws = null;
        SimpleTheme = null;
        wbp = null;
        this.DocumentProperties = null;

        dictAutoFitSharedStringCache = null;
        dictAutoFitTextCache = null;
        dictAutoFitFontCache = null;
        dictAutoFitFormatCodeCache = null;
        dictAutoFitTextRotationCache = null;

        dictBuiltInNumberingFormat = null;
        dictBuiltInNumberingFormatHash = null;

        dictStyleNumberingFormat = null;
        dictStyleNumberingFormatHash = null;

        //countStyle = 0;
        listStyle = null;
        dictStyleHash = null;

        //countStyleFont = 0;
        listStyleFont = null;
        dictStyleFontHash = null;

        //countStyleFill = 0;
        listStyleFill = null;
        dictStyleFillHash = null;

        //countStyleBorder = 0;
        listStyleBorder = null;
        dictStyleBorderHash = null;

        //countStyleCellStyle = 0;
        listStyleCellStyle = null;
        dictStyleCellStyleHash = null;

        //countStyleCellStyleFormat = 0;
        listStyleCellStyleFormat = null;
        dictStyleCellStyleFormatHash = null;

        //countStyleDifferentialFormat = 0;
        listStyleDifferentialFormat = null;
        dictStyleDifferentialFormatHash = null;

        //countStyleTableStyle = 0;
        listStyleTableStyle = null;
        dictStyleTableStyleHash = null;

        //countSharedString = 0;
        listSharedString = null;
        dictSharedStringHash = null;
        hsUniqueSharedString = null;

        StylesheetColors = null;
        TableStylesDefaultTableStyle = null;
        TableStylesDefaultPivotStyle = null;
    }

    /// <summary>
    /// Saves the spreadsheet. If it's a newly created spreadsheet, the default blank file name is used. If it's an existing spreadsheet, the given file name is used. WARNING: The existing spreadsheet will be overwritten without prompts.
    /// NOTE: This is a terminal method, meaning all internal memory stores will be nulled, cleared and nuked after executing this method. If you want some data, get it before this.
    /// </summary>
    public void Save()
    {
        CloseAndCleanUp();

        byte[] data = memstream.ToArray();
        memstream.Close();
        File.WriteAllBytes(gsSpreadsheetFileName, data);
    }

    /// <summary>
    /// Saves the spreadsheet to a given file name.
    /// NOTE: This is a terminal method, meaning all internal memory stores will be nulled, cleared and nuked after executing this method. If you want some data, get it before this.
    /// </summary>
    /// <param name="FileName">The file name of the spreadsheet to be saved to.</param>
    public void SaveAs(string FileName)
    {
        //gsSpreadsheetFileName = FileName;

        CloseAndCleanUp();

        byte[] data = memstream.ToArray();
        memstream.Close();
        //File.WriteAllBytes(gsSpreadsheetFileName, data);
        File.WriteAllBytes(FileName, data);
    }

    /// <summary>
    /// Saves the spreadsheet to a stream.
    /// NOTE: This is a terminal method, meaning all internal memory stores will be nulled, cleared and nuked after executing this method. If you want some data, get it before this.
    /// </summary>
    /// <param name="OutputStream">The output stream.</param>
    public void SaveAs(Stream OutputStream)
    {
        CloseAndCleanUp();

        memstream.WriteTo(OutputStream);
        memstream.Close();
    }

    /// <summary>
    /// Close the spreadsheet without saving.
    /// NOTE: This is a terminal method, meaning all internal memory stores will be nulled, cleared and nuked after executing this method. If you want some data, get it before this.
    /// </summary>
    public void CloseWithoutSaving()
    {
        CloseAndCleanUp();

        memstream.Close();
    }

    /// <summary>
    /// Releases all resources.
    /// </summary>
    public void Dispose()
    {
        NullifyInternalDataStores();

        xl.Dispose();
        memstream.Dispose();
    }
}
