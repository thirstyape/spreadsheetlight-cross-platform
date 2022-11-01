using System;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing
{
    /// <summary>
    /// Encapsulates properties and methods for specifying reflection effects.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Reflection class.
    /// </summary>
    public class SLReflection
    {
        internal bool HasReflection
        {
            get
            {
                return this.decBlurRadius != 0 || this.decStartOpacity != 100 || this.decStartPosition != 0
                    || this.decEndAlpha != 0 || this.decEndPosition != 100 || this.decDistance != 0
                    || this.decDirection != 0 || this.decFadeDirection != 90
                    || this.decHorizontalRatio != 100 || this.decVerticalRatio != 100
                    || this.decHorizontalSkew != 0 || this.decVerticalSkew != 0
                    || this.Alignment != A.RectangleAlignmentValues.Bottom || !this.RotateWithShape;
            }
        }

        internal decimal decBlurRadius;
        /// <summary>
        /// Blur radius of the reflection, ranging from 0 pt to 2147483647 pt. A suggested range is 0 pt to 100 pt. Accurate to 1/12700 of a point. Default value is 0 pt.
        /// </summary>
        public decimal BlurRadius
        {
            get { return decBlurRadius; }
            set
            {
                decBlurRadius = value;
                if (decBlurRadius < 0m) decBlurRadius = 0m;
                if (decBlurRadius > 2147483647m) decBlurRadius = 2147483647m;
            }
        }

        internal decimal decStartOpacity;
        /// <summary>
        /// Start opacity of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 100%.
        /// </summary>
        public decimal StartOpacity
        {
            get { return decStartOpacity; }
            set
            {
                decStartOpacity = value;
                if (decStartOpacity < 0m) decStartOpacity = 0m;
                if (decStartOpacity > 100m) decStartOpacity = 100m;
            }
        }

        internal decimal decStartPosition;
        /// <summary>
        /// Position of start opacity of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 0%.
        /// </summary>
        public decimal StartPosition
        {
            get { return decStartPosition; }
            set
            {
                decStartPosition = value;
                if (decStartPosition < 0m) decStartPosition = 0m;
                if (decStartPosition > 100m) decStartPosition = 100m;
            }
        }

        internal decimal decEndAlpha;
        /// <summary>
        /// End alpha of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 0%.
        /// </summary>
        public decimal EndAlpha
        {
            get { return decEndAlpha; }
            set
            {
                decEndAlpha = value;
                if (decEndAlpha < 0m) decEndAlpha = 0m;
                if (decEndAlpha > 100m) decEndAlpha = 100m;
            }
        }

        internal decimal decEndPosition;
        /// <summary>
        /// Position of end alpha of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 100%.
        /// </summary>
        public decimal EndPosition
        {
            get { return decEndPosition; }
            set
            {
                decEndPosition = value;
                if (decEndPosition < 0m) decEndPosition = 0m;
                if (decEndPosition > 100m) decEndPosition = 100m;
            }
        }

        internal decimal decDistance;
        /// <summary>
        /// Distance of the reflection from the origin, ranging from 0 pt to 2147483647 pt. A suggested range is 0 pt to 100 pt. Accurate to 1/12700 of a point. Default value is 0 pt.
        /// </summary>
        public decimal Distance
        {
            get { return decDistance; }
            set
            {
                decDistance = value;
                if (decDistance < 0m) decDistance = 0m;
                if (decDistance > 2147483647m) decDistance = 2147483647m;
            }
        }

        internal decimal decDirection;
        /// <summary>
        /// Direction of the alpha gradient, ranging from 0 degrees to 359.9 degrees. 0 degrees means to the right, 90 degrees is below, 180 degrees is to the right, and 270 degrees is above. Accurate to 1/60000 of a degree. Default value is 0 degrees.
        /// </summary>
        public decimal Direction
        {
            get { return decDirection; }
            set
            {
                decDirection = value;
                if (decDirection < 0m) decDirection = 0m;
                if (decDirection >= 360m) decDirection = 359.9m;
            }
        }

        internal decimal decFadeDirection;
        /// <summary>
        /// Direction to fade the reflection, ranging from 0 degrees to 359.9 degrees. 0 degrees means to the right, 90 degrees is below, 180 degrees is to the right, and 270 degrees is above. Accurate to 1/60000 of a degree. Default value is 90 degrees.
        /// </summary>
        public decimal FadeDirection
        {
            get { return decFadeDirection; }
            set
            {
                decFadeDirection = value;
                if (decFadeDirection < 0m) decFadeDirection = 0m;
                if (decFadeDirection >= 360m) decFadeDirection = 359.9m;
            }
        }

        internal decimal decHorizontalRatio;
        /// <summary>
        /// Horizontal scaling ratio in percentage. A negative ratio flips the reflection horizontally. Accurate to 1/1000 of a percent. Default value is 100%.
        /// </summary>
        public decimal HorizontalRatio
        {
            get { return decHorizontalRatio; }
            set { decHorizontalRatio = value; }
        }

        internal decimal decVerticalRatio;
        /// <summary>
        /// Vertical scaling ratio in percentage. A negative ratio flips the reflection vertically. Accurate to 1/1000 of a percent. Default value is 100%.
        /// </summary>
        public decimal VerticalRatio
        {
            get { return decVerticalRatio; }
            set { decVerticalRatio = value; }
        }

        internal decimal decHorizontalSkew;
        /// <summary>
        /// Horizontal skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.
        /// </summary>
        public decimal HorizontalSkew
        {
            get { return decHorizontalSkew; }
            set
            {
                decHorizontalSkew = value;
                if (decHorizontalSkew < -90m) decHorizontalSkew = -90m;
                if (decHorizontalSkew > 90m) decHorizontalSkew = 90m;
            }
        }

        internal decimal decVerticalSkew;
        /// <summary>
        /// Vertical skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.
        /// </summary>
        public decimal VerticalSkew
        {
            get { return decVerticalSkew; }
            set
            {
                decVerticalSkew = value;
                if (decVerticalSkew < -90m) decVerticalSkew = -90m;
                if (decVerticalSkew > 90m) decVerticalSkew = 90m;
            }
        }

        /// <summary>
        /// Sets the origin for the size scaling, angle skews and distance offsets. Default value is Bottom.
        /// </summary>
        public A.RectangleAlignmentValues Alignment { get; set; }

        /// <summary>
        /// True if the reflection should rotate as well. False otherwise.
        /// </summary>
        public bool RotateWithShape { get; set; }

        /// <summary>
        /// Initializes an instance of SLReflection.
        /// </summary>
        public SLReflection()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.decBlurRadius = 0;
            this.decStartOpacity = 100;
            this.decStartPosition = 0;
            this.decEndAlpha = 0;
            this.decEndPosition = 100;
            this.decDistance = 0;
            this.decDirection = 0;
            this.decFadeDirection = 90;
            this.decHorizontalRatio = 100;
            this.decVerticalRatio = 100;
            this.decHorizontalSkew = 0;
            this.decVerticalSkew = 0;
            this.Alignment = A.RectangleAlignmentValues.Bottom;
            this.RotateWithShape = true;
        }

        /// <summary>
        /// Set a tight reflection.
        /// </summary>
        public void SetTightReflection()
        {
            this.SetTightReflection(0m);
        }

        /// <summary>
        /// Set a tight reflection.
        /// </summary>
        /// <param name="Offset">Offset distance of the reflection, ranging from 0 pt to 2147483647 pt. A suggested range is 0 pt to 100 pt. Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        public void SetTightReflection(decimal Offset)
        {
            this.SetReflection(0.5m, 50m, 0m, 0.3m, 35m, Offset, 90m, 90m, 100m, -100m, 0m, 0m, A.RectangleAlignmentValues.BottomLeft, false);
        }

        /// <summary>
        /// Set a half reflection.
        /// </summary>
        public void SetHalfReflection()
        {
            this.SetHalfReflection(0m);
        }

        /// <summary>
        /// Set a half reflection.
        /// </summary>
        /// <param name="Offset">Offset distance of the reflection, ranging from 0 pt to 2147483647 pt. A suggested range is 0 pt to 100 pt. Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        public void SetHalfReflection(decimal Offset)
        {
            this.SetReflection(0.5m, 50m, 0m, 0.3m, 55m, Offset, 90m, 90m, 100m, -100m, 0m, 0m, A.RectangleAlignmentValues.BottomLeft, false);
        }

        /// <summary>
        /// Set a full reflection.
        /// </summary>
        public void SetFullReflection()
        {
            this.SetFullReflection(0m);
        }

        /// <summary>
        /// Set a full reflection.
        /// </summary>
        /// <param name="Offset">Offset distance of the reflection, ranging from 0 pt to 2147483647 pt. A suggested range is 0 pt to 100 pt. Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        public void SetFullReflection(decimal Offset)
        {
            this.SetReflection(0.5m, 50m, 0m, 0.3m, 90m, Offset, 90m, 90m, 100m, -100m, 0m, 0m, A.RectangleAlignmentValues.BottomLeft, false);
        }

        /// <summary>
        /// Set reflection.
        /// </summary>
        /// <param name="Transparency">Transparency ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Size">Size of reflection ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Distance">Distance of the reflection from the origin, ranging from 0 pt to 2147483647 pt. A suggested range is 0 pt to 100 pt. Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Blur">Blur radius of the reflection, ranging from 0 pt to 2147483647 pt. A suggested range is 0 pt to 100 pt. Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        public void SetReflection(decimal Transparency, decimal Size, decimal Distance, decimal Blur)
        {
            this.BlurRadius = Blur;
            this.StartOpacity = (100m - Transparency);
            this.EndAlpha = 0.3m;
            this.EndPosition = Size;
            this.Distance = Distance;
            this.VerticalRatio = -100;
            this.Alignment = A.RectangleAlignmentValues.BottomLeft;
            this.RotateWithShape = false;
        }

        /// <summary>
        /// Set a reflection of the picture.
        /// </summary>
        /// <param name="BlurRadius">Blur radius of the reflection, ranging from 0 pt to 2147483647 pt. A suggested range is 0 pt to 100 pt. Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="StartOpacity">Start opacity of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="StartPosition">Position of start opacity of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 0%.</param>
        /// <param name="EndAlpha">End alpha of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 0%.</param>
        /// <param name="EndPosition">Position of end alpha of the reflection, ranging from 0% to 100%. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="Distance">Distance of the reflection from the origin, ranging from 0 pt to 2147483647 pt. A suggested range is 0 pt to 100 pt. Accurate to 1/12700 of a point. Default value is 0 pt.</param>
        /// <param name="Direction">Direction of the alpha gradient relative to the origin, ranging from 0 degrees to 359.9 degrees. 0 degrees means to the right, 90 degrees is below, 180 degrees is to the right, and 270 degrees is above. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="FadeDirection">Direction to fade the reflection, ranging from 0 degrees to 359.9 degrees. 0 degrees means to the right, 90 degrees is below, 180 degrees is to the right, and 270 degrees is above. Accurate to 1/60000 of a degree. Default value is 90 degrees.</param>
        /// <param name="HorizontalRatio">Horizontal scaling ratio in percentage. A negative ratio flips the reflection horizontally. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="VerticalRatio">Vertical scaling ratio in percentage. A negative ratio flips the reflection vertically. Accurate to 1/1000 of a percent. Default value is 100%.</param>
        /// <param name="HorizontalSkew">Horizontal skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="VerticalSkew">Vertical skew angle, ranging from -90 degrees to 90 degrees. Accurate to 1/60000 of a degree. Default value is 0 degrees.</param>
        /// <param name="Alignment">Sets the origin for the size scaling, angle skews and distance offsets. Default value is Bottom.</param>
        /// <param name="RotateWithShape">True if the reflection should rotate. False otherwise. Default value is true.</param>
        public void SetReflection(decimal BlurRadius, decimal StartOpacity, decimal StartPosition, decimal EndAlpha, decimal EndPosition, decimal Distance, decimal Direction, decimal FadeDirection, decimal HorizontalRatio, decimal VerticalRatio, decimal HorizontalSkew, decimal VerticalSkew, A.RectangleAlignmentValues Alignment, bool RotateWithShape)
        {
            this.BlurRadius = BlurRadius;
            this.StartOpacity = StartOpacity;
            this.StartPosition = StartPosition;
            this.EndAlpha = EndAlpha;
            this.EndPosition = EndPosition;
            this.Distance = Distance;
            this.Direction = Direction;
            this.FadeDirection = FadeDirection;
            this.HorizontalRatio = HorizontalRatio;
            this.VerticalRatio = VerticalRatio;
            this.HorizontalSkew = HorizontalSkew;
            this.VerticalSkew = VerticalSkew;
            this.Alignment = Alignment;
            this.RotateWithShape = RotateWithShape;
        }

        internal A.Reflection ToReflection()
        {
            A.Reflection r = new A.Reflection();

            if (this.decBlurRadius != 0) r.BlurRadius = SLDrawingTool.CalculatePositiveCoordinate(this.decBlurRadius);
            if (this.decStartOpacity != 100) r.StartOpacity = SLDrawingTool.CalculatePositiveFixedPercentage(this.decStartOpacity);
            if (this.decStartPosition != 0) r.StartPosition = SLDrawingTool.CalculatePositiveFixedPercentage(this.decStartPosition);
            if (this.decEndAlpha != 0) r.EndAlpha = SLDrawingTool.CalculatePositiveFixedPercentage(this.decEndAlpha);
            if (this.decEndPosition != 100) r.EndPosition = SLDrawingTool.CalculatePositiveFixedPercentage(this.decEndPosition);
            if (this.decDistance != 0) r.Distance = SLDrawingTool.CalculatePositiveCoordinate(this.decDistance);
            if (this.decDirection != 0) r.Direction = SLDrawingTool.CalculatePositiveFixedAngle(this.decDirection);
            if (this.decFadeDirection != 90) r.FadeDirection = SLDrawingTool.CalculatePositiveFixedAngle(this.decFadeDirection);
            if (this.decHorizontalRatio != 100) r.HorizontalRatio = SLDrawingTool.CalculatePercentage(this.decHorizontalRatio);
            if (this.decVerticalRatio != 100) r.VerticalRatio = SLDrawingTool.CalculatePercentage(this.decVerticalRatio);
            if (this.decHorizontalSkew != 0) r.HorizontalSkew = SLDrawingTool.CalculateFixedAngle(this.decHorizontalSkew);
            if (this.decVerticalSkew != 0) r.VerticalSkew = SLDrawingTool.CalculateFixedAngle(this.decVerticalSkew);
            if (this.Alignment != A.RectangleAlignmentValues.Bottom) r.Alignment = this.Alignment;
            if (this.RotateWithShape != true) r.RotateWithShape = this.RotateWithShape;

            return r;
        }

        internal SLReflection Clone()
        {
            SLReflection r = new SLReflection();
            r.decBlurRadius = this.decBlurRadius;
            r.decStartOpacity = this.decStartOpacity;
            r.decStartPosition = this.decStartPosition;
            r.decEndAlpha = this.decEndAlpha;
            r.decEndPosition = this.decEndPosition;
            r.decDistance = this.decDistance;
            r.decDirection = this.decDirection;
            r.decFadeDirection = this.decFadeDirection;
            r.decHorizontalRatio = this.decHorizontalRatio;
            r.decVerticalRatio = this.decVerticalRatio;
            r.decHorizontalSkew = this.decHorizontalSkew;
            r.decVerticalSkew = this.decVerticalSkew;
            r.Alignment = this.Alignment;
            r.RotateWithShape = this.RotateWithShape;

            return r;
        }
    }
}
