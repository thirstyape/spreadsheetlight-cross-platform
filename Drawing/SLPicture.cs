using DocumentFormat.OpenXml.Packaging;
using System.Runtime.InteropServices;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing;

/// <summary>
/// Encapsulates properties and methods for a picture to be inserted into a worksheet.
/// </summary>
public class SLPicture
{
    // as opposed to data in byte array
    internal bool DataIsInFile;
    internal string PictureFileName;
    internal byte[] PictureByteData;
    internal ImagePartType PictureImagePartType = ImagePartType.Bmp;

    internal double TopPosition;
    internal double LeftPosition;
    internal bool UseEasyPositioning;

    // as opposed to absolute position. Not supporting TwoCellAnchor
    internal bool UseRelativePositioning;

    // used when relative positioning
    internal int AnchorRowIndex;
    internal int AnchorColumnIndex;

    // in units of EMU
    internal long OffsetX;
    internal long OffsetY;
    internal long WidthInEMU;
    internal long HeightInEMU;

    internal int WidthInPixels;
    internal int HeightInPixels;

    private float fHorizontalResolution;
    /// <summary>
    /// The horizontal resolution (DPI) of the picture. This is read-only.
    /// </summary>
    public float HorizontalResolution
    {
        get { return fHorizontalResolution; }
    }

    private float fVerticalResolution;
    /// <summary>
    /// The vertical resolution (DPI) of the picture. This is read-only.
    /// </summary>
    public float VerticalResolution
    {
        get { return fVerticalResolution; }
    }

    private float fTargetHorizontalResolution;
    private float fTargetVerticalResolution;
    private float fCurrentHorizontalResolution;
    private float fCurrentVerticalResolution;

    private float fHorizontalResolutionRatio;
    private float fVerticalResolutionRatio;

    private string sAlternativeText;
    /// <summary>
    /// The text used to assist users with disabilities. This is similar to the alt tag used in HTML.
    /// </summary>
    public string AlternativeText
    {
        get { return sAlternativeText; }
        set { sAlternativeText = value; }
    }

    private bool bLockWithSheet;
    /// <summary>
    /// Indicates whether the picture can be selected (selection is disabled when this is true). Locking the picture only works when the sheet is also protected. Default value is true.
    /// </summary>
    public bool LockWithSheet
    {
        get { return bLockWithSheet; }
        set { bLockWithSheet = value; }
    }

    private bool bPrintWithSheet;
    /// <summary>
    /// Indicates whether the picture is printed when the sheet is printed. Default value is true.
    /// </summary>
    public bool PrintWithSheet
    {
        get { return bPrintWithSheet; }
        set { bPrintWithSheet = value; }
    }

    private A.BlipCompressionValues vCompressionState;
    /// <summary>
    /// Compression settings for the picture. Default value is Print.
    /// </summary>
    public A.BlipCompressionValues CompressionState
    {
        get { return vCompressionState; }
        set { vCompressionState = value; }
    }

    private decimal decBrightness;
    /// <summary>
    /// Picture brightness modifier, ranging from -100% to 100%. Accurate to 1/1000 of a percent. Default value is 0%.
    /// </summary>
    public decimal Brightness
    {
        get { return decBrightness; }
        set
        {
            decBrightness = decimal.Round(value, 3);
            if (decBrightness < -100m) decBrightness = -100m;
            if (decBrightness > 100m) decBrightness = 100m;
        }
    }

    private decimal decContrast;
    /// <summary>
    /// Picture contrast modifier, ranging from -100% to 100%. Accurate to 1/1000 of a percent. Default value is 0%.
    /// </summary>
    public decimal Contrast
    {
        get { return decContrast; }
        set
        {
            decContrast = decimal.Round(value, 3);
            if (decContrast < -100m) decContrast = -100m;
            if (decContrast > 100m) decContrast = 100m;
        }
    }

    // not supporting yet because you need to change the positional offsets too.
    //private decimal decRotationAngle;
    ///// <summary>
    ///// The rotation angle in degrees, ranging from -3600 degrees to 3600 degrees. Accurate to 1/60000 of a degree. The angle increases clockwise, starting from the 12 o'clock position. Default value is 0 degrees.
    ///// </summary>
    //public decimal RotationAngle
    //{
    //    get { return decRotationAngle; }
    //    set
    //    {
    //        decRotationAngle = value;
    //        if (decRotationAngle < -3600m) decRotationAngle = -3600m;
    //        if (decRotationAngle > 3600m) decRotationAngle = 3600m;
    //    }
    //}

    internal SLShapeProperties ShapeProperties;

    /// <summary>
    /// Set the shape of the picture. Default value is Rectangle.
    /// </summary>
    public A.ShapeTypeValues PictureShape
    {
        get { return this.ShapeProperties.PresetGeometry; }
        set { this.ShapeProperties.PresetGeometry = value; }
    }

    /// <summary>
    /// Fill properties.
    /// </summary>
    public SLFill Fill { get { return this.ShapeProperties.Fill; } }

    /// <summary>
    /// Line properties.
    /// </summary>
    public SLLinePropertiesType Line { get { return this.ShapeProperties.Outline; } }

    /// <summary>
    /// Shadow properties.
    /// </summary>
    public SLShadowEffect Shadow { get { return this.ShapeProperties.EffectList.Shadow; } }

    /// <summary>
    /// Reflection properties.
    /// </summary>
    public SLReflection Reflection { get { return this.ShapeProperties.EffectList.Reflection; } }

    /// <summary>
    /// Glow properties.
    /// </summary>
    public SLGlow Glow { get { return this.ShapeProperties.EffectList.Glow; } }

    /// <summary>
    /// Soft edge properties.
    /// </summary>
    public SLSoftEdge SoftEdge { get { return this.ShapeProperties.EffectList.SoftEdge; } }

    /// <summary>
    /// 3D format properties.
    /// </summary>
    public SLFormat3D Format3D { get { return this.ShapeProperties.Format3D; } }

    /// <summary>
    /// 3D rotation properties.
    /// </summary>
    public SLRotation3D Rotation3D { get { return this.ShapeProperties.Rotation3D; } }

    internal bool HasUri;
    internal string HyperlinkUri;
    internal System.UriKind HyperlinkUriKind;
    internal bool IsHyperlinkExternal;

    internal SLPicture()
    {
    }

    /// <summary>
    /// Initializes an instance of SLPicture given the file name of a picture.
    /// </summary>
    /// <param name="PictureFileName">The file name of a picture to be inserted.</param>
    public SLPicture(string PictureFileName)
    {
        InitialisePicture(false);

        DataIsInFile = true;
        InitialisePictureFile(PictureFileName);

        SetResolution(false, 96, 96);
    }

    /// <summary>
    /// Initializes an instance of SLPicture given the file name of a picture.
    /// </summary>
    /// <param name="PictureFileName">The file name of a picture to be inserted.</param>
    /// <param name="ThrowExceptionsIfAny">Set to true to bubble exceptions up if there are any occurring within SpreadsheetLight. Set to false otherwise. The default is false.</param>
    public SLPicture(string PictureFileName, bool ThrowExceptionsIfAny)
    {
        InitialisePicture(ThrowExceptionsIfAny);

        DataIsInFile = true;
        InitialisePictureFile(PictureFileName);

        SetResolution(false, 96, 96);
    }

    /// <summary>
    /// Initializes an instance of SLPicture given the file name of a picture, and the targeted computer's horizontal and vertical resolution. This scales the picture according to how it will be displayed on the targeted computer screen.
    /// </summary>
    /// <param name="PictureFileName">The file name of a picture to be inserted.</param>
    /// <param name="TargetHorizontalResolution">The targeted computer's horizontal resolution (DPI).</param>
    /// <param name="TargetVerticalResolution">The targeted computer's vertical resolution (DPI).</param>
    public SLPicture(string PictureFileName, float TargetHorizontalResolution, float TargetVerticalResolution)
    {
        InitialisePicture(false);

        DataIsInFile = true;
        InitialisePictureFile(PictureFileName);

        SetResolution(true, TargetHorizontalResolution, TargetVerticalResolution);
    }

    /// <summary>
    /// Initializes an instance of SLPicture given the file name of a picture, and the targeted computer's horizontal and vertical resolution. This scales the picture according to how it will be displayed on the targeted computer screen.
    /// </summary>
    /// <param name="PictureFileName">The file name of a picture to be inserted.</param>
    /// <param name="TargetHorizontalResolution">The targeted computer's horizontal resolution (DPI).</param>
    /// <param name="TargetVerticalResolution">The targeted computer's vertical resolution (DPI).</param>
    /// <param name="ThrowExceptionsIfAny">Set to true to bubble exceptions up if there are any occurring within SpreadsheetLight. Set to false otherwise. The default is false.</param>
    public SLPicture(string PictureFileName, float TargetHorizontalResolution, float TargetVerticalResolution, bool ThrowExceptionsIfAny)
    {
        InitialisePicture(ThrowExceptionsIfAny);

        DataIsInFile = true;
        InitialisePictureFile(PictureFileName);

        SetResolution(true, TargetHorizontalResolution, TargetVerticalResolution);
    }

    // byte array as picture data suggested by Rob Hutchinson, with sample code sent in.

    /// <summary>
    /// Initializes an instance of SLPicture given a picture's data in a byte array.
    /// </summary>
    /// <param name="PictureByteData">The picture's data in a byte array.</param>
    /// <param name="PictureType">The image type of the picture.</param>
    public SLPicture(byte[] PictureByteData, ImagePartType PictureType)
    {
        InitialisePicture(false);

        DataIsInFile = false;
        this.PictureByteData = PictureByteData;
        this.PictureImagePartType = PictureType;

        SetResolution(false, 96, 96);
    }

    /// <summary>
    /// Initializes an instance of SLPicture given a picture's data in a byte array.
    /// </summary>
    /// <param name="PictureByteData">The picture's data in a byte array.</param>
    /// <param name="PictureType">The image type of the picture.</param>
    /// <param name="ThrowExceptionsIfAny">Set to true to bubble exceptions up if there are any occurring within SpreadsheetLight. Set to false otherwise. The default is false.</param>
    public SLPicture(byte[] PictureByteData, ImagePartType PictureType, bool ThrowExceptionsIfAny)
    {
        InitialisePicture(ThrowExceptionsIfAny);

        DataIsInFile = false;
        this.PictureByteData = PictureByteData;
        this.PictureImagePartType = PictureType;

        SetResolution(false, 96, 96);
    }

    /// <summary>
    /// Initializes an instance of SLPicture given a picture's data in a byte array, and the targeted computer's horizontal and vertical resolution. This scales the picture according to how it will be displayed on the targeted computer screen.
    /// </summary>
    /// <param name="PictureByteData">The picture's data in a byte array.</param>
    /// <param name="PictureType">The image type of the picture.</param>
    /// <param name="TargetHorizontalResolution">The targeted computer's horizontal resolution (DPI).</param>
    /// <param name="TargetVerticalResolution">The targeted computer's vertical resolution (DPI).</param>
    public SLPicture(byte[] PictureByteData, ImagePartType PictureType, float TargetHorizontalResolution, float TargetVerticalResolution)
    {
        InitialisePicture(false);

        DataIsInFile = false;
        this.PictureByteData = new byte[PictureByteData.Length];
        for (int i = 0; i < PictureByteData.Length; ++i)
        {
            this.PictureByteData[i] = PictureByteData[i];
        }
        this.PictureImagePartType = PictureType;

        SetResolution(true, TargetHorizontalResolution, TargetVerticalResolution);
    }

    /// <summary>
    /// Initializes an instance of SLPicture given a picture's data in a byte array, and the targeted computer's horizontal and vertical resolution. This scales the picture according to how it will be displayed on the targeted computer screen.
    /// </summary>
    /// <param name="PictureByteData">The picture's data in a byte array.</param>
    /// <param name="PictureType">The image type of the picture.</param>
    /// <param name="TargetHorizontalResolution">The targeted computer's horizontal resolution (DPI).</param>
    /// <param name="TargetVerticalResolution">The targeted computer's vertical resolution (DPI).</param>
    /// <param name="ThrowExceptionsIfAny">Set to true to bubble exceptions up if there are any occurring within SpreadsheetLight. Set to false otherwise. The default is false.</param>
    public SLPicture(byte[] PictureByteData, ImagePartType PictureType, float TargetHorizontalResolution, float TargetVerticalResolution, bool ThrowExceptionsIfAny)
    {
        InitialisePicture(ThrowExceptionsIfAny);

        DataIsInFile = false;
        this.PictureByteData = new byte[PictureByteData.Length];
        for (int i = 0; i < PictureByteData.Length; ++i)
        {
            this.PictureByteData[i] = PictureByteData[i];
        }
        this.PictureImagePartType = PictureType;

        SetResolution(true, TargetHorizontalResolution, TargetVerticalResolution);
    }

    private void InitialisePicture(bool ThrowExceptionsIfAny)
    {
        // should be true once we get *everyone* to stop using those confoundedly
        // hard to understand EMUs and absolute positionings...
        UseEasyPositioning = false;
        TopPosition = 0;
        LeftPosition = 0;

        UseRelativePositioning = true;
        AnchorRowIndex = 1;
        AnchorColumnIndex = 1;
        OffsetX = 0;
        OffsetY = 0;
        WidthInEMU = 0;
        HeightInEMU = 0;
        WidthInPixels = 0;
        HeightInPixels = 0;
        fHorizontalResolutionRatio = 1;
        fVerticalResolutionRatio = 1;

        this.bLockWithSheet = true;
        this.bPrintWithSheet = true;
        this.vCompressionState = A.BlipCompressionValues.Print;
        this.decBrightness = 0;
        this.decContrast = 0;
        //this.decRotationAngle = 0;

        this.ShapeProperties = new SLShapeProperties(new List<System.Drawing.Color>(), ThrowExceptionsIfAny);

        this.HasUri = false;
        this.HyperlinkUri = string.Empty;
        this.HyperlinkUriKind = UriKind.Absolute;
        this.IsHyperlinkExternal = true;

        this.DataIsInFile = true;
        this.PictureFileName = string.Empty;
        this.PictureByteData = new byte[1];
        this.PictureImagePartType = ImagePartType.Bmp;
    }

    private void InitialisePictureFile(string FileName)
    {
        this.PictureFileName = FileName.Trim();

        this.PictureImagePartType = SLDrawingTool.GetImagePartType(this.PictureFileName);

        string sImageFileName = this.PictureFileName.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
        sImageFileName = sImageFileName.Substring(sImageFileName.LastIndexOf(Path.DirectorySeparatorChar) + 1);
        this.sAlternativeText = sImageFileName;
    }

    private void SetResolution(bool HasTarget, float TargetHorizontalResolution, float TargetVerticalResolution)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) == false)
            return;

#pragma warning disable CA1416
        // this is used to sort of get the current computer's screen DPI
        System.Drawing.Bitmap bmResolution = new System.Drawing.Bitmap(32, 32);

        // thanks to Stefano Lanzavecchia for suggesting the use of System.Drawing.Image
        // as a general image loader as opposed to the Bitmap class.
        // This allows the use of EMF images (and other image types that the Image class
        // supports).
        System.Drawing.Image img;
        if (this.DataIsInFile)
        {
            img = System.Drawing.Image.FromFile(this.PictureFileName);
        }
        else
        {
            using (MemoryStream ms = new MemoryStream(this.PictureByteData))
            {
                img = System.Drawing.Image.FromStream(ms);
            }
        }

        this.fHorizontalResolution = img.HorizontalResolution;
        this.fVerticalResolution = img.VerticalResolution;

        if (HasTarget)
        {
            this.fTargetHorizontalResolution = TargetHorizontalResolution;
            this.fTargetVerticalResolution = TargetVerticalResolution;
        }
        else
        {
            this.fTargetHorizontalResolution = bmResolution.HorizontalResolution;
            this.fTargetVerticalResolution = bmResolution.VerticalResolution;
        }

        this.fCurrentHorizontalResolution = bmResolution.HorizontalResolution;
        this.fCurrentVerticalResolution = bmResolution.VerticalResolution;
        this.fHorizontalResolutionRatio = this.fTargetHorizontalResolution / this.fCurrentHorizontalResolution;
        this.fVerticalResolutionRatio = this.fTargetVerticalResolution / this.fCurrentVerticalResolution;

        this.WidthInPixels = img.Width;
        this.HeightInPixels = img.Height;
        this.ResizeInPixels(img.Width, img.Height);
        img.Dispose();
        bmResolution.Dispose();
#pragma warning restore CA1416
    }

    /// <summary>
    /// Set the position of the picture relative to the top-left of the worksheet.
    /// </summary>
    /// <param name="Top">Top position based on row index. For example, 0.5 means at the half-way point of the 1st row, 2.5 means at the half-way point of the 3rd row.</param>
    /// <param name="Left">Left position based on column index. For example, 0.5 means at the half-way point of the 1st column, 2.5 means at the half-way point of the 3rd column.</param>
    public void SetPosition(double Top, double Left)
    {
        // make sure to do the calculation upon insertion
        this.UseEasyPositioning = true;
        this.TopPosition = Top;
        this.LeftPosition = Left;
        this.UseRelativePositioning = true;
        this.OffsetX = 0;
        this.OffsetY = 0;
    }

    /// <summary>
    /// Resize the picture with new size dimensions using percentages of the original size dimensions.
    /// </summary>
    /// <param name="ScaleWidth">A percentage of the original width. It is suggested to keep the range between 1% and 56624%.</param>
    /// <param name="ScaleHeight">A percentage of the original height. It is suggested to keep the range between 1% and 56624%.</param>
    public void ResizeInPercentage(int ScaleWidth, int ScaleHeight)
    {
        int iNewWidth = Convert.ToInt32((decimal)this.WidthInPixels * (decimal)ScaleWidth / 100m);
        int iNewHeight = Convert.ToInt32((decimal)this.HeightInPixels * (decimal)ScaleHeight / 100m);
        this.ResizeInPixels(iNewWidth, iNewHeight);
    }

    /// <summary>
    /// Resize the picture with new size dimensions in pixels.
    /// </summary>
    /// <param name="Width">The new width in pixels.</param>
    /// <param name="Height">The new height in pixels.</param>
    public void ResizeInPixels(int Width, int Height)
    {
        long lWidthInEMU = Convert.ToInt64((float)Width * this.fHorizontalResolutionRatio * (float)SLConstants.InchToEMU / this.HorizontalResolution);
        long lHeightInEMU = Convert.ToInt64((float)Height * this.fVerticalResolutionRatio * (float)SLConstants.InchToEMU / this.VerticalResolution);
        this.ResizeInEMU(lWidthInEMU, lHeightInEMU);
    }

    /// <summary>
    /// Resize the picture with new size dimension in English Metric Units (EMUs).
    /// </summary>
    /// <param name="Width">The new width in EMUs.</param>
    /// <param name="Height">The new height in EMUs.</param>
    public void ResizeInEMU(long Width, long Height)
    {
        this.WidthInEMU = Width;
        this.HeightInEMU = Height;
    }

    /// <summary>
    /// Inserts a hyperlink to a webpage.
    /// </summary>
    /// <param name="URL">The target webpage URL.</param>
    public void InsertUrlHyperlink(string URL)
    {
        this.HasUri = true;
        this.HyperlinkUri = URL;
        this.HyperlinkUriKind = UriKind.Absolute;
        this.IsHyperlinkExternal = true;
    }

    /// <summary>
    /// Inserts a hyperlink to a document on the computer.
    /// </summary>
    /// <param name="FilePath">The relative path to the file based on the location of the spreadsheet.</param>
    public void InsertFileHyperlink(string FilePath)
    {
        this.HasUri = true;
        this.HyperlinkUri = FilePath;
        this.HyperlinkUriKind = UriKind.Relative;
        this.IsHyperlinkExternal = true;
    }

    /// <summary>
    /// Inserts a hyperlink to an email address.
    /// </summary>
    /// <param name="EmailAddress">The email address, such as johndoe@acme.com</param>
    public void InsertEmailHyperlink(string EmailAddress)
    {
        this.HasUri = true;
        this.HyperlinkUri = string.Format("mailto:{0}", EmailAddress);
        this.HyperlinkUriKind = UriKind.Absolute;
        this.IsHyperlinkExternal = true;
    }

    /// <summary>
    /// Inserts a hyperlink to a place in the spreadsheet document.
    /// </summary>
    /// <param name="SheetName">The name of the worksheet being referenced.</param>
    /// <param name="RowIndex">The row index of the referenced cell. If this is invalid, row 1 will be used.</param>
    /// <param name="ColumnIndex">The column index of the referenced cell. If this is invalid, column 1 will be used.</param>
    public void InsertInternalHyperlink(string SheetName, int RowIndex, int ColumnIndex)
    {
        int iRowIndex = RowIndex;
        int iColumnIndex = ColumnIndex;
        if (iRowIndex < 1 || iRowIndex > SLConstants.RowLimit) iRowIndex = 1;
        if (iColumnIndex < 1 || iColumnIndex > SLConstants.ColumnLimit) iColumnIndex = 1;

        this.HasUri = true;
        this.HyperlinkUri = string.Format("#{0}!{1}", SLTool.FormatWorksheetNameForFormula(SheetName), SLTool.ToCellReference(iRowIndex, iColumnIndex));
        this.HyperlinkUriKind = UriKind.Relative;
        this.IsHyperlinkExternal = false;
    }

    /// <summary>
    /// Inserts a hyperlink to a place in the spreadsheet document.
    /// </summary>
    /// <param name="SheetName">The name of the worksheet being referenced.</param>
    /// <param name="CellReference">The cell reference, such as "A1".</param>
    public void InsertInternalHyperlink(string SheetName, string CellReference)
    {
        this.HasUri = true;
        this.HyperlinkUri = string.Format("#{0}!{1}", SLTool.FormatWorksheetNameForFormula(SheetName), CellReference);
        this.HyperlinkUriKind = UriKind.Relative;
        this.IsHyperlinkExternal = false;
    }

    /// <summary>
    /// Inserts a hyperlink to a place in the spreadsheet document.
    /// </summary>
    /// <param name="DefinedName">A defined name in the spreadsheet.</param>
    public void InsertInternalHyperlink(string DefinedName)
    {
        this.HasUri = true;
        this.HyperlinkUri = string.Format("#{0}", DefinedName);
        this.HyperlinkUriKind = UriKind.Relative;
        this.IsHyperlinkExternal = false;
    }

    internal SLPicture Clone()
    {
        SLPicture pic = new SLPicture();
        pic.DataIsInFile = this.DataIsInFile;
        pic.PictureFileName = this.PictureFileName;
        pic.PictureByteData = new byte[this.PictureByteData.Length];
        for (int i = 0; i < this.PictureByteData.Length; ++i)
        {
            pic.PictureByteData[i] = this.PictureByteData[i];
        }
        pic.PictureImagePartType = this.PictureImagePartType;

        pic.TopPosition = this.TopPosition;
        pic.LeftPosition = this.LeftPosition;
        pic.UseEasyPositioning = this.UseEasyPositioning;
        pic.UseRelativePositioning = this.UseRelativePositioning;
        pic.AnchorRowIndex = this.AnchorRowIndex;
        pic.AnchorColumnIndex = this.AnchorColumnIndex;
        pic.OffsetX = this.OffsetX;
        pic.OffsetY = this.OffsetY;
        pic.WidthInEMU = this.WidthInEMU;
        pic.HeightInEMU = this.HeightInEMU;
        pic.WidthInPixels = this.WidthInPixels;
        pic.HeightInPixels = this.HeightInPixels;
        pic.fHorizontalResolution = this.fHorizontalResolution;
        pic.fVerticalResolution = this.fVerticalResolution;
        pic.fTargetHorizontalResolution = this.fTargetHorizontalResolution;
        pic.fTargetVerticalResolution = this.fTargetVerticalResolution;
        pic.fCurrentHorizontalResolution = this.fCurrentHorizontalResolution;
        pic.fCurrentVerticalResolution = this.fCurrentVerticalResolution;
        pic.fHorizontalResolutionRatio = this.fHorizontalResolutionRatio;
        pic.fVerticalResolutionRatio = this.fVerticalResolutionRatio;
        pic.sAlternativeText = this.sAlternativeText;
        pic.bLockWithSheet = this.bLockWithSheet;
        pic.bPrintWithSheet = this.bPrintWithSheet;
        pic.vCompressionState = this.vCompressionState;
        pic.decBrightness = this.decBrightness;
        pic.decContrast = this.decContrast;

        pic.ShapeProperties = this.ShapeProperties.Clone();

        pic.HasUri = this.HasUri;
        pic.HyperlinkUri = this.HyperlinkUri;
        pic.HyperlinkUriKind = this.HyperlinkUriKind;
        pic.IsHyperlinkExternal = this.IsHyperlinkExternal;

        return pic;
    }
}
