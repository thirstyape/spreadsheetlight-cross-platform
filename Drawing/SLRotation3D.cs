using System;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing
{
    /// <summary>
    /// Specifies camera preset settings.
    /// </summary>
    public enum SLCameraPresetValues
    {
        /// <summary>
        /// None
        /// </summary>
        None = 0,
        /// <summary>
        /// Isometric Left Down
        /// </summary>
        IsometricLeftDown,
        /// <summary>
        /// Isometric Right Up
        /// </summary>
        IsometricRightUp,
        /// <summary>
        /// Isometric Top Up
        /// </summary>
        IsometricTopUp,
        /// <summary>
        /// Isometric Bottom Down
        /// </summary>
        IsometricBottomDown,
        /// <summary>
        /// Off Axis 1 Left
        /// </summary>
        OffAxis1Left,
        /// <summary>
        /// Off Axis 1 Right
        /// </summary>
        OffAxis1Right,
        /// <summary>
        /// Off Axis 1 Top
        /// </summary>
        OffAxis1Top,
        /// <summary>
        /// Off Axis 2 Left
        /// </summary>
        OffAxis2Left,
        /// <summary>
        /// Off Axis 2 Right
        /// </summary>
        OffAxis2Right,
        /// <summary>
        /// Off Axis 2 Top
        /// </summary>
        OffAxis2Top,
        /// <summary>
        /// Perspective Front
        /// </summary>
        PerspectiveFront,
        /// <summary>
        /// Perspective Left
        /// </summary>
        PerspectiveLeft,
        /// <summary>
        /// Perspective Right
        /// </summary>
        PerspectiveRight,
        /// <summary>
        /// Perspective Below
        /// </summary>
        PerspectiveBelow,
        /// <summary>
        /// Perspective Above
        /// </summary>
        PerspectiveAbove,
        /// <summary>
        /// Perspective Relaxed Moderately
        /// </summary>
        PerspectiveRelaxedModerately,
        /// <summary>
        /// Perspective Relaxed
        /// </summary>
        PerspectiveRelaxed,
        /// <summary>
        /// Perspective Contrasting Left
        /// </summary>
        PerspectiveContrastingLeft,
        /// <summary>
        /// Perspective Contrasting Right
        /// </summary>
        PerspectiveContrastingRight,
        /// <summary>
        /// Perspective Heroic Extreme Left
        /// </summary>
        PerspectiveHeroicExtremeLeft,
        /// <summary>
        /// Perspective Heroic Extreme Right
        /// </summary>
        PerspectiveHeroicExtremeRight,
        /// <summary>
        /// Oblique Top Left
        /// </summary>
        ObliqueTopLeft,
        /// <summary>
        /// Oblique Top Right
        /// </summary>
        ObliqueTopRight,
        /// <summary>
        /// Oblique Bottom Left
        /// </summary>
        ObliqueBottomLeft,
        /// <summary>
        /// Oblique Bottom Right
        /// </summary>
        ObliqueBottomRight
    }

    /// <summary>
    /// Encapsulates 3D rotation properties. Works together with SLFormat3D class.
    /// This simulates some properties of DocumentFormat.OpenXml.Drawing.Scene3DType
    /// and DocumentFormat.OpenXml.Drawing.Shape3DType classes. The reason for this mixing
    /// is because Excel separates different properties from both classes into 2 separate sections
    /// on the user interface (3-D Format and 3-D Rotation). Hence SLRotation3D and SLFormat3D
    /// classes instead of straightforward mapping of the SDK Scene3DType and Shape3DType classes.
    /// </summary>
    public class SLRotation3D
    {
        internal bool HasCamera;

        internal A.PresetCameraValues CameraPreset { get; set; }

        internal bool HasXYZSet;
        internal bool HasPerspectiveSet;

        internal decimal decX;
        /// <summary>
        /// Longitude angle ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.
        /// </summary>
        public decimal X
        {
            get { return decX; }
            set
            {
                this.decX = value;
                if (this.decX < 0m) this.decX = 0m;
                if (this.decX >= 360m) this.decX = 359.9m;
                this.HasCamera = true;
                this.HasXYZSet = true;
            }
        }

        internal decimal decY;
        /// <summary>
        /// Latitude angle ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.
        /// </summary>
        public decimal Y
        {
            get { return decY; }
            set
            {
                this.decY = value;
                if (this.decY < 0m) this.decY = 0m;
                if (this.decY >= 360m) this.decY = 359.9m;
                this.HasCamera = true;
                this.HasXYZSet = true;
            }
        }

        internal decimal decZ;
        /// <summary>
        /// Revolution angle ranging from 0 degrees to 359.9 degrees. Accurate to 1/60000 of a degree.
        /// </summary>
        public decimal Z
        {
            get { return decZ; }
            set
            {
                this.decZ = value;
                if (this.decZ < 0m) this.decZ = 0m;
                if (this.decZ >= 360m) this.decZ = 359.9m;
                this.HasCamera = true;
                this.HasXYZSet = true;
            }
        }

        internal decimal decPerspective;
        /// <summary>
        /// Perspective angle ranging from 0 degrees to 180 degrees. However, a suggested maximum is 120 degrees.
        /// </summary>
        public decimal Perspective
        {
            get { return decPerspective; }
            set
            {
                if (this.IsPerspectiveView(this.CameraPreset))
                {
                    this.decPerspective = value;
                    if (this.decPerspective < 0m) this.decPerspective = 0m;
                    if (this.decPerspective > 180m) this.decPerspective = 180m;
                    this.HasCamera = true;
                    this.HasPerspectiveSet = true;
                }
            }
        }

        internal decimal decDistanceZ;
        /// <summary>
        /// Distance from the ground, ranging from -2147483648 pt to 2147483647 pt. However, a suggested range is -4000 pt to 4000 pt.
        /// </summary>
        public decimal DistanceZ
        {
            get { return decDistanceZ; }
            set
            {
                this.decDistanceZ = value;
                if (this.decDistanceZ < -2147483648m) this.decDistanceZ = -2147483648m;
                if (this.decDistanceZ > 2147483647m) this.decDistanceZ = 2147483647m;
            }
        }

        /// <summary>
        /// Initializes an instance of SLRotation3D.
        /// </summary>
        public SLRotation3D()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.HasCamera = false;
            this.HasXYZSet = false;
            this.HasPerspectiveSet = false;
            this.CameraPreset = A.PresetCameraValues.OrthographicFront;
            this.decX = 0;
            this.decY = 0;
            this.decZ = 0;
            this.decPerspective = 0;
            this.decDistanceZ = 0;
        }

        /// <summary>
        /// Set camera settings using a preset.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        public void SetCameraPreset(SLCameraPresetValues Preset)
        {
            switch (Preset)
            {
                case SLCameraPresetValues.None:
                    this.CameraPreset = A.PresetCameraValues.OrthographicFront;
                    this.decX = 0;
                    this.decY = 0;
                    this.decZ = 0;
                    this.decPerspective = 0;
                    this.HasCamera = false;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.IsometricLeftDown:
                    this.CameraPreset = A.PresetCameraValues.IsometricLeftDown;
                    this.decX = 45;
                    this.decY = 35;
                    this.decZ = 0;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.IsometricRightUp:
                    this.CameraPreset = A.PresetCameraValues.IsometricRightUp;
                    this.decX = 315;
                    this.decY = 35;
                    this.decZ = 0;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.IsometricTopUp:
                    this.CameraPreset = A.PresetCameraValues.IsometricTopUp;
                    this.decX = 314.7m;
                    this.decY = 324.6m;
                    this.decZ = 60.2m;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.IsometricBottomDown:
                    this.CameraPreset = A.PresetCameraValues.IsometricBottomDown;
                    this.decX = 314.7m;
                    this.decY = 35.39999999999999m;
                    this.decZ = 299.8m;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis1Left:
                    this.CameraPreset = A.PresetCameraValues.IsometricOffAxis1Left;
                    this.decX = 64m;
                    this.decY = 18m;
                    this.decZ = 0;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis1Right:
                    this.CameraPreset = A.PresetCameraValues.IsometricOffAxis1Right;
                    this.decX = 334m;
                    this.decY = 18m;
                    this.decZ = 0;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis1Top:
                    this.CameraPreset = A.PresetCameraValues.IsometricOffAxis1Top;
                    this.decX = 306.5m;
                    this.decY = 301.3m;
                    this.decZ = 57.6m;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis2Left:
                    this.CameraPreset = A.PresetCameraValues.IsometricOffAxis2Left;
                    this.decX = 26m;
                    this.decY = 18m;
                    this.decZ = 0m;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis2Right:
                    this.CameraPreset = A.PresetCameraValues.IsometricOffAxis2Right;
                    this.decX = 296m;
                    this.decY = 18m;
                    this.decZ = 0m;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.OffAxis2Top:
                    this.CameraPreset = A.PresetCameraValues.IsometricOffAxis2Top;
                    this.decX = 53.49999999999999m;
                    this.decY = 301.3m;
                    this.decZ = 302.4m;
                    this.decPerspective = 0;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveFront:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveFront;
                    this.decX = 0m;
                    this.decY = 0m;
                    this.decZ = 0m;
                    this.decPerspective = 45m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveLeft:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveLeft;
                    this.decX = 20m;
                    this.decY = 0m;
                    this.decZ = 0m;
                    this.decPerspective = 45m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveRight:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveRight;
                    this.decX = 340m;
                    this.decY = 0m;
                    this.decZ = 0m;
                    this.decPerspective = 45m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveBelow:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveBelow;
                    this.decX = 0m;
                    this.decY = 20m;
                    this.decZ = 0m;
                    this.decPerspective = 45m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveAbove:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveAbove;
                    this.decX = 0m;
                    this.decY = 340m;
                    this.decZ = 0m;
                    this.decPerspective = 45m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveRelaxedModerately:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveRelaxedModerately;
                    this.decX = 0m;
                    this.decY = 324.8m;
                    this.decZ = 0m;
                    this.decPerspective = 45m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveRelaxed:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveRelaxed;
                    this.decX = 0m;
                    this.decY = 309.6m;
                    this.decZ = 0m;
                    this.decPerspective = 45m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveContrastingLeft:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveContrastingLeftFacing;
                    this.decX = 43.89999999999999m;
                    this.decY = 10.4m;
                    this.decZ = 356.4m;
                    this.decPerspective = 45m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveContrastingRight:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveContrastingRightFacing;
                    this.decX = 316.1m;
                    this.decY = 10.4m;
                    this.decZ = 3.6m;
                    this.decPerspective = 45m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveHeroicExtremeLeft:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveHeroicExtremeLeftFacing;
                    this.decX = 34.49999999999999m;
                    this.decY = 8.1m;
                    this.decZ = 357.1m;
                    this.decPerspective = 80m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.PerspectiveHeroicExtremeRight:
                    this.CameraPreset = A.PresetCameraValues.PerspectiveHeroicExtremeRightFacing;
                    this.decX = 325.5m;
                    this.decY = 8.1m;
                    this.decZ = 2.9m;
                    this.decPerspective = 80m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.ObliqueTopLeft:
                    this.CameraPreset = A.PresetCameraValues.ObliqueTopLeft;
                    this.decX = 0m;
                    this.decY = 0m;
                    this.decZ = 0m;
                    this.decPerspective = 0m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.ObliqueTopRight:
                    this.CameraPreset = A.PresetCameraValues.ObliqueTopRight;
                    this.decX = 0m;
                    this.decY = 0m;
                    this.decZ = 0m;
                    this.decPerspective = 0m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.ObliqueBottomLeft:
                    this.CameraPreset = A.PresetCameraValues.ObliqueBottomLeft;
                    this.decX = 0m;
                    this.decY = 0m;
                    this.decZ = 0m;
                    this.decPerspective = 0m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
                case SLCameraPresetValues.ObliqueBottomRight:
                    this.CameraPreset = A.PresetCameraValues.ObliqueBottomRight;
                    this.decX = 0m;
                    this.decY = 0m;
                    this.decZ = 0m;
                    this.decPerspective = 0m;
                    this.HasCamera = true;
                    this.HasXYZSet = false;
                    this.HasPerspectiveSet = false;
                    break;
            }
        }

        private bool IsPerspectiveView(A.PresetCameraValues Preset)
        {
            bool result = false;

            switch (Preset)
            {
                case A.PresetCameraValues.LegacyPerspectiveBottom:
                case A.PresetCameraValues.LegacyPerspectiveBottomLeft:
                case A.PresetCameraValues.LegacyPerspectiveBottomRight:
                case A.PresetCameraValues.LegacyPerspectiveFront:
                case A.PresetCameraValues.LegacyPerspectiveLeft:
                case A.PresetCameraValues.LegacyPerspectiveRight:
                case A.PresetCameraValues.LegacyPerspectiveTop:
                case A.PresetCameraValues.LegacyPerspectiveTopLeft:
                case A.PresetCameraValues.LegacyPerspectiveTopRight:
                case A.PresetCameraValues.PerspectiveAbove:
                case A.PresetCameraValues.PerspectiveAboveLeftFacing:
                case A.PresetCameraValues.PerspectiveAboveRightFacing:
                case A.PresetCameraValues.PerspectiveBelow:
                case A.PresetCameraValues.PerspectiveContrastingLeftFacing:
                case A.PresetCameraValues.PerspectiveContrastingRightFacing:
                case A.PresetCameraValues.PerspectiveFront:
                case A.PresetCameraValues.PerspectiveHeroicExtremeLeftFacing:
                case A.PresetCameraValues.PerspectiveHeroicExtremeRightFacing:
                case A.PresetCameraValues.PerspectiveHeroicLeftFacing:
                case A.PresetCameraValues.PerspectiveHeroicRightFacing:
                case A.PresetCameraValues.PerspectiveLeft:
                case A.PresetCameraValues.PerspectiveRelaxed:
                case A.PresetCameraValues.PerspectiveRelaxedModerately:
                case A.PresetCameraValues.PerspectiveRight:
                    result = true;
                    break;
            }

            return result;
        }

        internal SLRotation3D Clone()
        {
            SLRotation3D rot = new SLRotation3D();
            rot.HasCamera = this.HasCamera;
            rot.CameraPreset = this.CameraPreset;
            rot.HasXYZSet = this.HasXYZSet;
            rot.HasPerspectiveSet = this.HasPerspectiveSet;
            rot.decX = this.decX;
            rot.decY = this.decY;
            rot.decZ = this.decZ;
            rot.decPerspective = this.decPerspective;
            rot.decDistanceZ = this.decDistanceZ;

            return rot;
        }
    }
}
