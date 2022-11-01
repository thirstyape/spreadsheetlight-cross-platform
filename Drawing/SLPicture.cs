using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing
{
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
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the absolute position of the picture in pixels relative to the top-left corner of the worksheet.
        /// Consider using the SetPosition() function instead.
        /// </summary>
        /// <param name="OffsetX">Offset from the left of the worksheet in pixels.</param>
        /// <param name="OffsetY">Offset from the top of the worksheet in pixels.</param>
        [Obsolete("This is an esoteric function. Use the easier-to-understand SetPosition() function instead.")]
        public void SetAbsolutePositionInPixels(int OffsetX, int OffsetY)
        {
            // absolute position is influenced by the image resolution
            long lOffsetXinEMU = Convert.ToInt64((float)OffsetX * this.fHorizontalResolutionRatio * (float)SLConstants.InchToEMU / this.HorizontalResolution);
            long lOffsetYinEMU = Convert.ToInt64((float)OffsetY * this.fVerticalResolutionRatio * (float)SLConstants.InchToEMU / this.VerticalResolution);
            //this.SetAbsolutePositionInEMU(lOffsetXinEMU, lOffsetYinEMU);

            this.UseEasyPositioning = false;
            this.UseRelativePositioning = false;
            this.OffsetX = lOffsetXinEMU;
            this.OffsetY = lOffsetYinEMU;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the absolute position of the picture in English Metric Units (EMUs) relative to the top-left corner of the worksheet.
        /// Consider using the SetPosition() function instead.
        /// </summary>
        /// <param name="OffsetX">Offset from the left of the worksheet in EMUs.</param>
        /// <param name="OffsetY">Offset from the top of the worksheet in EMUs.</param>
        [Obsolete("This is an esoteric function. Use the easier-to-understand SetPosition() function instead.")]
        public void SetAbsolutePositionInEMU(long OffsetX, long OffsetY)
        {
            this.UseEasyPositioning = false;
            this.UseRelativePositioning = false;
            this.OffsetX = OffsetX;
            this.OffsetY = OffsetY;
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
        /// <strong>Obsolete. </strong>Set the position of the picture in pixels relative to the top-left of the worksheet. The picture is anchored to the top-left corner of a given cell.
        /// Consider using the SetPosition() function instead.
        /// </summary>
        /// <param name="AnchorRowIndex">Row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">Column index of the anchor cell.</param>
        /// <param name="OffsetX">Offset from the left of the anchor cell in pixels.</param>
        /// <param name="OffsetY">Offset from the top of the anchor cell in pixels.</param>
        [Obsolete("This is an esoteric function. Use the easier-to-understand SetPosition() function instead.")]
        public void SetRelativePositionInPixels(int AnchorRowIndex, int AnchorColumnIndex, int OffsetX, int OffsetY)
        {
            long lOffsetXinEMU = (long)OffsetX * SLDocument.PixelToEMU;
            long lOffsetYinEMU = (long)OffsetY * SLDocument.PixelToEMU;
            //this.SetRelativePositionInEMU(AnchorRowIndex, AnchorColumnIndex, lOffsetXinEMU, lOffsetYinEMU);

            this.UseEasyPositioning = false;
            this.UseRelativePositioning = true;
            this.OffsetX = lOffsetXinEMU;
            this.OffsetY = lOffsetYinEMU;

            this.AnchorRowIndex = AnchorRowIndex;
            if (this.AnchorRowIndex < 1) this.AnchorRowIndex = 1;
            if (this.AnchorRowIndex > SLConstants.RowLimit) this.AnchorRowIndex = SLConstants.RowLimit;

            this.AnchorColumnIndex = AnchorColumnIndex;
            if (this.AnchorColumnIndex < 1) this.AnchorColumnIndex = 1;
            if (this.AnchorColumnIndex > SLConstants.ColumnLimit) this.AnchorColumnIndex = SLConstants.ColumnLimit;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the position of the picture in English Metric Units (EMUs) relative to the top-left of the worksheet. The picture is anchored to the top-left corner of a given cell.
        /// Consider using the SetPosition() function instead.
        /// </summary>
        /// <param name="AnchorRowIndex">Row index of the anchor cell.</param>
        /// <param name="AnchorColumnIndex">Column index of the anchor cell.</param>
        /// <param name="OffsetX">Offset from the left of the anchor cell in EMUs.</param>
        /// <param name="OffsetY">Offset from the top of the anchor cell in EMUs.</param>
        [Obsolete("This is an esoteric function. Use the easier-to-understand SetPosition() function instead.")]
        public void SetRelativePositionInEMU(int AnchorRowIndex, int AnchorColumnIndex, long OffsetX, long OffsetY)
        {
            this.UseEasyPositioning = false;
            this.UseRelativePositioning = true;
            this.OffsetX = OffsetX;
            this.OffsetY = OffsetY;

            this.AnchorRowIndex = AnchorRowIndex;
            if (this.AnchorRowIndex < 1) this.AnchorRowIndex = 1;
            if (this.AnchorRowIndex > SLConstants.RowLimit) this.AnchorRowIndex = SLConstants.RowLimit;

            this.AnchorColumnIndex = AnchorColumnIndex;
            if (this.AnchorColumnIndex < 1) this.AnchorColumnIndex = 1;
            if (this.AnchorColumnIndex > SLConstants.ColumnLimit) this.AnchorColumnIndex = SLConstants.ColumnLimit;
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
        /// <strong>Obsolete. </strong>Fill the background of the picture with color. The color will be seen through the transparent parts of the picture.
        /// </summary>
        /// <param name="FillColor">The color used to fill the background of the picture.</param>
        /// <param name="Transparency">Transparency of the fill color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        [Obsolete("Use the Fill property instead.")]
        public void SetSolidFill(System.Drawing.Color FillColor, decimal Transparency)
        {
            this.Fill.SetSolidFill(FillColor, Transparency);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Fill the background of the picture with color. The color will be seen through the transparent parts of the picture.
        /// </summary>
        /// <param name="ThemeColor">The theme color used to fill the background of the picture.</param>
        /// <param name="Transparency">Transparency of the fill color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        [Obsolete("Use the Fill property instead.")]
        public void SetSolidFill(A.SchemeColorValues ThemeColor, decimal Transparency)
        {
            this.Fill.SetSolidFill(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), 0, Transparency);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Fill the background of the picture with color. The color will be seen through the transparent parts of the picture.
        /// </summary>
        /// <param name="ThemeColor">The theme color used to fill the background of the picture.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the fill color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        [Obsolete("Use the Fill property instead.")]
        public void SetSolidFill(A.SchemeColorValues ThemeColor, decimal Tint, decimal Transparency)
        {
            this.Fill.SetSolidFill(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), (double)Tint, Transparency);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the outline color.
        /// </summary>
        /// <param name="OutlineColor">The color used to outline the picture.</param>
        /// <param name="Transparency">Transparency of the outline color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        [Obsolete("Use the Line property instead.")]
        public void SetSolidOutline(System.Drawing.Color OutlineColor, decimal Transparency)
        {
            this.Line.SetSolidLine(OutlineColor, Transparency);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the outline color.
        /// </summary>
        /// <param name="ThemeColor">The theme color used to outline the picture.</param>
        /// <param name="Transparency">Transparency of the outline color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        [Obsolete("Use the Line property instead.")]
        public void SetSolidOutline(A.SchemeColorValues ThemeColor, decimal Transparency)
        {
            this.Line.SetSolidLine(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), 0, Transparency);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the outline color.
        /// </summary>
        /// <param name="ThemeColor">The theme color used to outline the picture.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the outline color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        [Obsolete("Use the Line property instead.")]
        public void SetSolidOutline(A.SchemeColorValues ThemeColor, decimal Tint, decimal Transparency)
        {
            this.Line.SetSolidLine(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), (double)Tint, Transparency);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the outline style of the picture.
        /// </summary>
        /// <param name="Width">Width of the outline, between 0 pt and 1584 pt. Accurate to 1/12700 of a point.</param>
        /// <param name="CompoundType">Compound type. Default value is single.</param>
        /// <param name="DashType">Dash style of the outline.</param>
        /// <param name="CapType">Line cap type of the outline. Default value is square.</param>
        /// <param name="JoinType">Join type of the outline at the corners. Default value is round.</param>
        [Obsolete("Use the Line property instead.")]
        public void SetOutlineStyle(decimal Width, A.CompoundLineValues? CompoundType, A.PresetLineDashValues? DashType, A.LineCapValues? CapType, SLLineJoinValues? JoinType)
        {
            this.Line.Width = Width;
            if (CompoundType != null) this.Line.CompoundLineType = CompoundType.Value;
            if (DashType != null) this.Line.DashType = DashType.Value;
            if (CapType != null) this.Line.CapType = CapType.Value;
            if (JoinType != null) this.Line.JoinType = JoinType.Value;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set an inner shadow of the picture.
        /// </summary>
        /// <param name="ShadowColor">The color used for the inner shadow.</param>
        /// <param name="Transparency">Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Blur">Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Angle">Angle of shadow projection based on picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Distance">Distance of shadow away from picture, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Shadow property instead.")]
        public void SetInnerShadow(System.Drawing.Color ShadowColor, decimal Transparency, decimal Blur, decimal Angle, decimal Distance)
        {
            this.Shadow.IsInnerShadow = true;
            this.Shadow.SetShadowColor(ShadowColor, 0);
            this.Shadow.Transparency = Transparency;
            this.Shadow.InnerShadowBlurRadius = Blur;
            this.Shadow.Angle = Angle;
            this.Shadow.InnerShadowDistance = Distance;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set an inner shadow of the picture.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the inner shadow.</param>
        /// <param name="Transparency">Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Blur">Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Angle">Angle of shadow projection based on picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Distance">Distance of shadow away from picture, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Shadow property instead.")]
        public void SetInnerShadow(A.SchemeColorValues ThemeColor, decimal Transparency, decimal Blur, decimal Angle, decimal Distance)
        {
            this.Shadow.IsInnerShadow = true;
            this.Shadow.SetShadowColor(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), 0, Transparency);
            this.Shadow.Transparency = Transparency;
            this.Shadow.InnerShadowBlurRadius = Blur;
            this.Shadow.Angle = Angle;
            this.Shadow.InnerShadowDistance = Distance;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set an inner shadow of the picture.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the inner shadow.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Blur">Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Angle">Angle of shadow projection based on picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Distance">Distance of shadow away from picture, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Shadow property instead.")]
        public void SetInnerShadow(A.SchemeColorValues ThemeColor, decimal Tint, decimal Transparency, decimal Blur, decimal Angle, decimal Distance)
        {
            this.Shadow.IsInnerShadow = true;
            this.Shadow.SetShadowColor(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), (double)Tint, Transparency);
            this.Shadow.Transparency = Transparency;
            this.Shadow.InnerShadowBlurRadius = Blur;
            this.Shadow.Angle = Angle;
            this.Shadow.InnerShadowDistance = Distance;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set an outer shadow of the picture.
        /// </summary>
        /// <param name="ShadowColor">The color used for the outer shadow.</param>
        /// <param name="Transparency">Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Size">Scale size of shadow based on size of picture in percentage (consider a range of 1% to 200%). Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="Blur">Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Angle">Angle of shadow projection based on picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Distance">Distance of shadow away from picture, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Alignment">Sets the origin of the picture for the size scaling. Default value is Bottom.</param>
        /// <param name="RotateWithPicture">True if the shadow should rotate with the picture if the picture is rotated. False otherwise. Default value is true.</param>
        [Obsolete("Use the Shadow property instead.")]
        public void SetOuterShadow(System.Drawing.Color ShadowColor, decimal Transparency, decimal Size, decimal Blur, decimal Angle, decimal Distance, A.RectangleAlignmentValues Alignment, bool RotateWithPicture)
        {
            this.Shadow.IsInnerShadow = false;
            this.Shadow.SetShadowColor(ShadowColor, Transparency);
            this.Shadow.Transparency = Transparency;
            this.Shadow.Size = Size;
            this.Shadow.OuterShadowBlurRadius = Blur;
            this.Shadow.Angle = Angle;
            this.Shadow.OuterShadowDistance = Distance;
            this.Shadow.OuterShadowAlignment = Alignment;
            this.Shadow.OuterShadowRotateWithShape = RotateWithPicture;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set an outer shadow of the picture.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the outer shadow.</param>
        /// <param name="Transparency">Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Size">Scale size of shadow based on size of picture in percentage (consider a range of 1% to 200%). Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="Blur">Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Angle">Angle of shadow projection based on picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Distance">Distance of shadow away from picture, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Alignment">Sets the origin of the picture for the size scaling. Default value is Bottom.</param>
        /// <param name="RotateWithPicture">True if the shadow should rotate with the picture if the picture is rotated. False otherwise. Default value is true.</param>
        [Obsolete("Use the Shadow property instead.")]
        public void SetOuterShadow(A.SchemeColorValues ThemeColor, decimal Transparency, decimal Size, decimal Blur, decimal Angle, decimal Distance, A.RectangleAlignmentValues Alignment, bool RotateWithPicture)
        {
            this.Shadow.IsInnerShadow = false;
            this.Shadow.SetShadowColor(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), 0, Transparency);
            this.Shadow.Transparency = Transparency;
            this.Shadow.Size = Size;
            this.Shadow.OuterShadowBlurRadius = Blur;
            this.Shadow.Angle = Angle;
            this.Shadow.OuterShadowDistance = Distance;
            this.Shadow.OuterShadowAlignment = Alignment;
            this.Shadow.OuterShadowRotateWithShape = RotateWithPicture;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set an outer shadow of the picture.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the outer shadow.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Size">Scale size of shadow based on size of picture in percentage (consider a range of 1% to 200%). Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="Blur">Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Angle">Angle of shadow projection based on picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Distance">Distance of shadow away from picture, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Alignment">Sets the origin of the picture for the size scaling. Default value is Bottom.</param>
        /// <param name="RotateWithPicture">True if the shadow should rotate with the picture if the picture is rotated. False otherwise. Default value is true.</param>
        [Obsolete("Use the Shadow property instead.")]
        public void SetOuterShadow(A.SchemeColorValues ThemeColor, decimal Tint, decimal Transparency, decimal Size, decimal Blur, decimal Angle, decimal Distance, A.RectangleAlignmentValues Alignment, bool RotateWithPicture)
        {
            this.Shadow.IsInnerShadow = false;
            this.Shadow.SetShadowColor(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), (double)Tint, Transparency);
            this.Shadow.Transparency = Transparency;
            this.Shadow.Size = Size;
            this.Shadow.OuterShadowBlurRadius = Blur;
            this.Shadow.Angle = Angle;
            this.Shadow.OuterShadowDistance = Distance;
            this.Shadow.OuterShadowAlignment = Alignment;
            this.Shadow.OuterShadowRotateWithShape = RotateWithPicture;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a perspective shadow of the picture.
        /// </summary>
        /// <param name="ShadowColor">The color used for the perspective shadow.</param>
        /// <param name="Transparency">Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="HorizontalRatio">Horizontal scaling ratio in percentage (consider a range of -200% to 200%). A negative ratio flips the shadow horizontally. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="VerticalRatio">Vertical scaling ratio in percentage (consider a range of -200% to 200%). A negative ratio flips the shadow vertically. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="HorizontalSkew">Horizontal skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="VerticalSkew">Vertical skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Blur">Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Angle">Angle of shadow projection based on picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Distance">Distance of shadow away from picture, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Alignment">Sets the origin of the picture for the size scaling, angle skews and distance offsets. Default value is Bottom.</param>
        /// <param name="RotateWithPicture">True if the shadow should rotate with the picture if the picture is rotated. False otherwise. Default value is true.</param>
        [Obsolete("Use the Shadow property instead.")]
        public void SetPerspectiveShadow(System.Drawing.Color ShadowColor, decimal Transparency, decimal HorizontalRatio, decimal VerticalRatio, decimal HorizontalSkew, decimal VerticalSkew, decimal Blur, decimal Angle, decimal Distance, A.RectangleAlignmentValues Alignment, bool RotateWithPicture)
        {
            this.Shadow.IsInnerShadow = false;
            this.Shadow.SetShadowColor(ShadowColor, Transparency);
            this.Shadow.Transparency = Transparency;
            this.Shadow.OuterShadowHorizontalRatio = HorizontalRatio;
            this.Shadow.OuterShadowVerticalRatio = VerticalRatio;
            this.Shadow.OuterShadowHorizontalSkew = HorizontalSkew;
            this.Shadow.OuterShadowVerticalSkew = VerticalSkew;
            this.Shadow.OuterShadowBlurRadius = Blur;
            this.Shadow.Angle = Angle;
            this.Shadow.OuterShadowDistance = Distance;
            this.Shadow.OuterShadowAlignment = Alignment;
            this.Shadow.OuterShadowRotateWithShape = RotateWithPicture;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a perspective shadow of the picture.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the perspective shadow.</param>
        /// <param name="Transparency">Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="HorizontalRatio">Horizontal scaling ratio in percentage (consider a range of -200% to 200%). A negative ratio flips the shadow horizontally. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="VerticalRatio">Vertical scaling ratio in percentage (consider a range of -200% to 200%). A negative ratio flips the shadow vertically. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="HorizontalSkew">Horizontal skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="VerticalSkew">Vertical skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Blur">Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Angle">Angle of shadow projection based on picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Distance">Distance of shadow away from picture, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Alignment">Sets the origin of the picture for the size scaling, angle skews and distance offsets. Default value is Bottom.</param>
        /// <param name="RotateWithPicture">True if the shadow should rotate with the picture if the picture is rotated. False otherwise. Default value is true.</param>
        [Obsolete("Use the Shadow property instead.")]
        public void SetPerspectiveShadow(A.SchemeColorValues ThemeColor, decimal Transparency, decimal HorizontalRatio, decimal VerticalRatio, decimal HorizontalSkew, decimal VerticalSkew, decimal Blur, decimal Angle, decimal Distance, A.RectangleAlignmentValues Alignment, bool RotateWithPicture)
        {
            this.Shadow.IsInnerShadow = false;
            this.Shadow.SetShadowColor(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), 0, Transparency);
            this.Shadow.Transparency = Transparency;
            this.Shadow.OuterShadowHorizontalRatio = HorizontalRatio;
            this.Shadow.OuterShadowVerticalRatio = VerticalRatio;
            this.Shadow.OuterShadowHorizontalSkew = HorizontalSkew;
            this.Shadow.OuterShadowVerticalSkew = VerticalSkew;
            this.Shadow.OuterShadowBlurRadius = Blur;
            this.Shadow.Angle = Angle;
            this.Shadow.OuterShadowDistance = Distance;
            this.Shadow.OuterShadowAlignment = Alignment;
            this.Shadow.OuterShadowRotateWithShape = RotateWithPicture;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a perspective shadow of the picture.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the perspective shadow.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="HorizontalRatio">Horizontal scaling ratio in percentage (consider a range of -200% to 200%). A negative ratio flips the shadow horizontally. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="VerticalRatio">Vertical scaling ratio in percentage (consider a range of -200% to 200%). A negative ratio flips the shadow vertically. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="HorizontalSkew">Horizontal skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="VerticalSkew">Vertical skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Blur">Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Angle">Angle of shadow projection based on picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Distance">Distance of shadow away from picture, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Alignment">Sets the origin of the picture for the size scaling, angle skews and distance offsets. Default value is Bottom.</param>
        /// <param name="RotateWithPicture">True if the shadow should rotate with the picture if the picture is rotated. False otherwise. Default value is true.</param>
        [Obsolete("Use the Shadow property instead.")]
        public void SetPerspectiveShadow(A.SchemeColorValues ThemeColor, decimal Tint, decimal Transparency, decimal HorizontalRatio, decimal VerticalRatio, decimal HorizontalSkew, decimal VerticalSkew, decimal Blur, decimal Angle, decimal Distance, A.RectangleAlignmentValues Alignment, bool RotateWithPicture)
        {
            this.Shadow.IsInnerShadow = false;
            this.Shadow.SetShadowColor(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), (double)Tint, Transparency);
            this.Shadow.Transparency = Transparency;
            this.Shadow.OuterShadowHorizontalRatio = HorizontalRatio;
            this.Shadow.OuterShadowVerticalRatio = VerticalRatio;
            this.Shadow.OuterShadowHorizontalSkew = HorizontalSkew;
            this.Shadow.OuterShadowVerticalSkew = VerticalSkew;
            this.Shadow.OuterShadowBlurRadius = Blur;
            this.Shadow.Angle = Angle;
            this.Shadow.OuterShadowDistance = Distance;
            this.Shadow.OuterShadowAlignment = Alignment;
            this.Shadow.OuterShadowRotateWithShape = RotateWithPicture;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a bevelled top.
        /// </summary>
        /// <param name="BevelPreset">The bevel type. Default value is circle.</param>
        /// <param name="Width">Width of the bevel, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 6 pt.</param>
        /// <param name="Height">Height of the bevel, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 6 pt.</param>
        [Obsolete("Use the Format3D property instead.")]
        public void Set3DBevelTop(A.BevelPresetValues BevelPreset, decimal Width, decimal Height)
        {
            this.Format3D.SetBevelTop(BevelPreset, Width, Height);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a bevelled bottom.
        /// </summary>
        /// <param name="BevelPreset">The bevel type. Default value is circle.</param>
        /// <param name="Width">Width of the bevel, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 6 pt.</param>
        /// <param name="Height">Height of the bevel, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 6 pt.</param>
        [Obsolete("Use the Format3D property instead.")]
        public void Set3DBevelBottom(A.BevelPresetValues BevelPreset, decimal Width, decimal Height)
        {
            this.Format3D.SetBevelBottom(BevelPreset, Width, Height);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the extrusion (or depth).
        /// </summary>
        /// <param name="ExtrusionColor">The color used for the extrusion.</param>
        /// <param name="Transparency">Transparency of the extrusion color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="ExtrusionHeight">Height of the extrusion, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Format3D property instead.")]
        public void Set3DExtrusion(System.Drawing.Color ExtrusionColor, decimal Transparency, decimal ExtrusionHeight)
        {
            this.Format3D.SetExtrusion(ExtrusionColor, ExtrusionHeight);
            this.Format3D.clrExtrusionColor.Transparency = Transparency;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the extrusion (or depth).
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the extrusion.</param>
        /// <param name="Transparency">Transparency of the theme color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="ExtrusionHeight">Height of the extrusion, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Format3D property instead.")]
        public void Set3DExtrusion(A.SchemeColorValues ThemeColor, decimal Transparency, decimal ExtrusionHeight)
        {
            this.Format3D.SetExtrusion(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), 0, ExtrusionHeight);
            this.Format3D.clrExtrusionColor.Transparency = Transparency;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the extrusion (or depth).
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the extrusion.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the theme color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="ExtrusionHeight">Height of the extrusion, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Format3D property instead.")]
        public void Set3DExtrusion(A.SchemeColorValues ThemeColor, decimal Tint, decimal Transparency, decimal ExtrusionHeight)
        {
            this.Format3D.SetExtrusion(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), (double)Tint, ExtrusionHeight);
            this.Format3D.clrExtrusionColor.Transparency = Transparency;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the 3D contour.
        /// </summary>
        /// <param name="ContourColor">The color used for the contour.</param>
        /// <param name="Transparency">Transparency of the contour color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="ContourWidth">Width of the contour, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Format3D property instead.")]
        public void Set3DContour(System.Drawing.Color ContourColor, decimal Transparency, decimal ContourWidth)
        {
            this.Format3D.SetContour(ContourColor, ContourWidth);
            this.Format3D.clrContourColor.Transparency = Transparency;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the 3D contour.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the contour.</param>
        /// <param name="Transparency">Transparency of the theme color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="ContourWidth">Width of the contour, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Format3D property instead.")]
        public void Set3DContour(A.SchemeColorValues ThemeColor, decimal Transparency, decimal ContourWidth)
        {
            this.Format3D.SetContour(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), 0, ContourWidth);
            this.Format3D.clrContourColor.Transparency = Transparency;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the 3D contour.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the contour.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the theme color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="ContourWidth">Width of the contour, ranging from 0 pt to 2147483647 pt (but consider a maximum of 1584 pt). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Format3D property instead.")]
        public void Set3DContour(A.SchemeColorValues ThemeColor, decimal Tint, decimal Transparency, decimal ContourWidth)
        {
            this.Format3D.SetContour(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), (double)Tint, ContourWidth);
            this.Format3D.clrContourColor.Transparency = Transparency;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the surface material.
        /// </summary>
        /// <param name="MaterialType">The material used. Default value is WarmMatte.</param>
        [Obsolete("Use the Format3D property instead.")]
        public void Set3DMaterialType(A.PresetMaterialTypeValues MaterialType)
        {
            this.Format3D.Material = MaterialType;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the Z distance.
        /// </summary>
        /// <param name="ZDistance">The Z distance, ranging from -4000 pt to 4000 pt. Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Rotation3D property instead.")]
        public void Set3DZDistance(decimal ZDistance)
        {
            this.Rotation3D.DistanceZ = ZDistance;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the camera and light properties.
        /// </summary>
        /// <param name="CameraPreset">A preset set of properties for the camera, which can be overridden. Default value is OrthographicFront.</param>
        /// <param name="FieldOfView">Field of view, ranging from 0 degrees to 180 degrees. Accurate to 1/60000 of a degree.</param>
        /// <param name="Zoom">Zoom percentage, ranging from 0% to 2147483.647%. Accurate to 1/1000 of a percent.</param>
        /// <param name="CameraLatitude">Camera latitude angle, ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.</param>
        /// <param name="CameraLongitude">Camera longitude angle, ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.</param>
        /// <param name="CameraRevolution">Camera revolution angle, ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.</param>
        /// <param name="LightRigType">The type of light used. Default value is ThreePoints.</param>
        /// <param name="LightRigDirection">The direction of the light. Default value is Top.</param>
        /// <param name="LightRigLatitude">Light rig latitude angle, ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.</param>
        /// <param name="LightRigLongitude">Light rig longitude angle, ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.</param>
        /// <param name="LightRigRevolution">Light rig revolution angle, ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.</param>
        /// <remarks>Imagine the screen to be the X-Y plane, the positive X-axis pointing to the right, and the positive Y-axis pointing up.
        /// The positive Z-axis points perpendicularly from the screen towards you.
        /// The latitude value increases as you turn around the X-axis, using the right-hand rule.
        /// The longitude value increases as you turn around the Y-axis, using the <em>left-hand rule</em> (meaning it decreases according to right-hand rule).
        /// The revolution value increases as you turn around the Z-axis, using the right-hand rule.
        /// And if you're mapping values directly from Microsoft Excel, don't treat the X, Y and Z values as values related to the axes.
        /// The latitude maps to the Y value, longitude maps to the X value, and revolution maps to the Z value.</remarks>
        [Obsolete("Use the Rotation3D and Format3D properties instead.")]
        public void Set3DScene(A.PresetCameraValues CameraPreset, decimal FieldOfView, decimal Zoom, decimal CameraLatitude, decimal CameraLongitude, decimal CameraRevolution, A.LightRigValues LightRigType, A.LightRigDirectionValues LightRigDirection, decimal LightRigLatitude, decimal LightRigLongitude, decimal LightRigRevolution)
        {
            this.Rotation3D.CameraPreset = CameraPreset;
            this.Rotation3D.Perspective = FieldOfView;
            // no zoom
            this.Rotation3D.Y = CameraLatitude;
            this.Rotation3D.X = CameraLongitude;
            this.Rotation3D.Z = CameraRevolution;
            this.Format3D.Lighting = LightRigType;
            // no light direction, latitude, longitude
            this.Format3D.Angle = LightRigRevolution;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set soft edges on the picture.
        /// </summary>
        /// <param name="Radius">Radius of the soft edge, ranging from 0 pt to 2147483647 pt (but consider a much lower maximum). Accurate to 1/12700 of a point.</param>
        [Obsolete("Use the SoftEdge property instead.")]
        public void SetSoftEdge(decimal Radius)
        {
            this.SoftEdge.Radius = Radius;
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the picture to glow on the edges.
        /// </summary>
        /// <param name="GlowColor">The color used for the glow.</param>
        /// <param name="Transparency">Transparency of the glow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="GlowRadius">Radius of the glow, ranging from 0 pt to 2147483647 pt (but consider a much lower maximum). Accurate to 1/12700 of a point.</param>
        [Obsolete("Use the Glow property instead.")]
        public void SetGlow(System.Drawing.Color GlowColor, decimal Transparency, decimal GlowRadius)
        {
            this.Glow.SetGlow(GlowColor, Transparency, GlowRadius);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the picture to glow on the edges.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the glow.</param>
        /// <param name="Transparency">Transparency of the theme color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="GlowRadius">Radius of the glow, ranging from 0 pt to 2147483647 pt (but consider a much lower maximum). Accurate to 1/12700 of a point.</param>
        [Obsolete("Use the Glow property instead.")]
        public void SetGlow(A.SchemeColorValues ThemeColor, decimal Transparency, decimal GlowRadius)
        {
            this.Glow.SetGlow(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), 0, Transparency, GlowRadius);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set the picture to glow on the edges.
        /// </summary>
        /// <param name="ThemeColor">The theme color used for the glow.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the theme color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="GlowRadius">Radius of the glow, ranging from 0 pt to 2147483647 pt (but consider a much lower maximum). Accurate to 1/12700 of a point.</param>
        [Obsolete("Use the Glow property instead.")]
        public void SetGlow(A.SchemeColorValues ThemeColor, decimal Tint, decimal Transparency, decimal GlowRadius)
        {
            this.Glow.SetGlow(SLDrawingTool.TranslateSchemeColorValue(ThemeColor), (double)Tint, Transparency, GlowRadius);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a tight reflection of the picture.
        /// </summary>
        [Obsolete("Use the Reflection property instead.")]
        public void SetTightReflection()
        {
            this.Reflection.SetTightReflection(0m);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a tight reflection of the picture.
        /// </summary>
        /// <param name="Offset">Offset distance of the reflection below the picture, ranging from 0 pt to 2147483647 pt (but consider a much lower maximum). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Reflection property instead.")]
        public void SetTightReflection(decimal Offset)
        {
            this.Reflection.SetReflection(0.5m, 50m, 0m, 0.3m, 35m, Offset, 90m, 90m, 100m, -100m, 0m, 0m, A.RectangleAlignmentValues.BottomLeft, false);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a reflection that's about half of the picture.
        /// </summary>
        [Obsolete("Use the Reflection property instead.")]
        public void SetHalfReflection()
        {
            this.Reflection.SetHalfReflection(0m);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a reflection that's about half of the picture.
        /// </summary>
        /// <param name="Offset">Offset distance of the reflection below the picture, ranging from 0 pt to 2147483647 pt (but consider a much lower maximum). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Reflection property instead.")]
        public void SetHalfReflection(decimal Offset)
        {
            this.Reflection.SetReflection(0.5m, 50m, 0m, 0.3m, 55m, Offset, 90m, 90m, 100m, -100m, 0m, 0m, A.RectangleAlignmentValues.BottomLeft, false);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a full reflection of the picture.
        /// </summary>
        [Obsolete("Use the Reflection property instead.")]
        public void SetFullReflection()
        {
            this.Reflection.SetFullReflection(0m);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a full reflection of the picture.
        /// </summary>
        /// <param name="Offset">Offset distance of the reflection below the picture, ranging from 0 pt to 2147483647 pt (but consider a much lower maximum). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        [Obsolete("Use the Reflection property instead.")]
        public void SetFullReflection(decimal Offset)
        {
            this.Reflection.SetReflection(0.5m, 50m, 0m, 0.3m, 90m, Offset, 90m, 90m, 100m, -100m, 0m, 0m, A.RectangleAlignmentValues.BottomLeft, false);
        }

        /// <summary>
        /// <strong>Obsolete. </strong>Set a reflection of the picture.
        /// </summary>
        /// <param name="BlurRadius">Blur radius of the reflection, ranging from 0 pt to 2147483647 pt (but consider a much lower maximum). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="StartOpacity">Start opacity of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="StartPosition">Position of start opacity of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 0%.</param>
        /// <param name="EndAlpha">End alpha of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 0%.</param>
        /// <param name="EndPosition">Position of end alpha of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="Distance">Distance of the reflection from the picture, ranging from 0 pt to 2147483647 pt (but consider a much lower maximum). Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Direction">Direction of the alpha gradient relative to the picture, ranging from 0 degrees to 359.9 degrees. 0 degrees means to the right, 90 degrees is below, 180 degrees is to the right, and 270 degrees is above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="FadeDirection">Direction to fade the reflection, ranging from 0 degrees to 359.9 degrees. 0 degrees means to the right, 90 degrees is below, 180 degrees is to the right, and 270 degrees is above. Accurate to 1/60000 of a degree. Default value is 90 degrees.</param>
        /// <param name="HorizontalRatio">Horizontal scaling ratio in percentage. A negative ratio flips the reflection horizontally. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="VerticalRatio">Vertical scaling ratio in percentage. A negative ratio flips the reflection vertically. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="HorizontalSkew">Horizontal skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="VerticalSkew">Vertical skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Alignment">Sets the origin of the picture for the size scaling, angle skews and distance offsets. Default value is Bottom.</param>
        /// <param name="RotateWithShape">True if the reflection should rotate with the picture if the picture is rotated. False otherwise. Default value is true.</param>
        [Obsolete("Use the Reflection property instead.")]
        public void SetReflection(decimal BlurRadius, decimal StartOpacity, decimal StartPosition, decimal EndAlpha, decimal EndPosition, decimal Distance, decimal Direction, decimal FadeDirection, decimal HorizontalRatio, decimal VerticalRatio, decimal HorizontalSkew, decimal VerticalSkew, A.RectangleAlignmentValues Alignment, bool RotateWithShape)
        {
            this.Reflection.SetReflection(BlurRadius, StartOpacity, StartPosition, EndAlpha, EndPosition, Distance, Direction, FadeDirection, HorizontalRatio, VerticalRatio, HorizontalSkew, VerticalSkew, Alignment, RotateWithShape);
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
}
