using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Drawing
{
    /// <summary>
    /// Encapsulates 3D shape properties. Works together with SLRotation3D class.
    /// This simulates some properties of DocumentFormat.OpenXml.Drawing.Scene3DType
    /// and DocumentFormat.OpenXml.Drawing.Shape3DType classes. The reason for this mixing
    /// is because Excel separates different properties from both classes into 2 separate sections
    /// on the user interface (3-D Format and 3-D Rotation). Hence SLRotation3D and SLFormat3D
    /// classes instead of straightforward mapping of the SDK Scene3DType and Shape3DType classes.
    /// </summary>
    public class SLFormat3D
    {
        internal List<System.Drawing.Color> listThemeColors;

        private bool bHasBevelTop;
        /// <summary>
        /// Specifies if there's a top bevel. This is read-only.
        /// </summary>
        public bool HasBevelTop { get { return this.bHasBevelTop; } }

        internal A.BevelPresetValues vBevelTopPreset;
        /// <summary>
        /// The bevel type of the top bevel. Default is circle.
        /// </summary>
        public A.BevelPresetValues BevelTopPreset
        {
            get { return this.vBevelTopPreset; }
            set
            {
                this.vBevelTopPreset = value;
                this.bHasBevelTop = true;
            }
        }

        internal decimal decBevelTopWidth;
        /// <summary>
        /// Width of the top bevel, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.
        /// </summary>
        public decimal BevelTopWidth
        {
            get { return this.decBevelTopWidth; }
            set
            {
                this.decBevelTopWidth = value;
                if (this.decBevelTopWidth < 0m) this.decBevelTopWidth = 0m;
                if (this.decBevelTopWidth > 2147483647m) this.decBevelTopWidth = 2147483647m;
                this.bHasBevelTop = true;
            }
        }

        internal decimal decBevelTopHeight;
        /// <summary>
        /// Height of the top bevel, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.
        /// </summary>
        public decimal BevelTopHeight
        {
            get { return this.decBevelTopHeight; }
            set
            {
                this.decBevelTopHeight = value;
                if (this.decBevelTopHeight < 0m) this.decBevelTopHeight = 0m;
                if (this.decBevelTopHeight > 2147483647m) this.decBevelTopHeight = 2147483647m;
                this.bHasBevelTop = true;
            }
        }

        private bool bHasBevelBottom;
        /// <summary>
        /// Specifies if there's a bottom bevel. This is read-only.
        /// </summary>
        public bool HasBevelBottom { get { return this.bHasBevelBottom; } }

        internal A.BevelPresetValues vBevelBottomPreset;

        /// <summary>
        /// The bevel type of the bottom bevel. Default is circle.
        /// </summary>
        public A.BevelPresetValues BevelBottomPreset
        {
            get { return this.vBevelBottomPreset; }
            set
            {
                this.vBevelBottomPreset = value;
                this.bHasBevelBottom = true;
            }
        }

        internal decimal decBevelBottomWidth;
        /// <summary>
        /// Width of the bottom bevel, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.
        /// </summary>
        public decimal BevelBottomWidth
        {
            get { return this.decBevelBottomWidth; }
            set
            {
                this.decBevelBottomWidth = value;
                if (this.decBevelBottomWidth < 0m) this.decBevelBottomWidth = 0m;
                if (this.decBevelBottomWidth > 2147483647m) this.decBevelBottomWidth = 2147483647m;
                this.bHasBevelBottom = true;
            }
        }

        internal decimal decBevelBottomHeight;
        /// <summary>
        /// Height of the bottom bevel, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.
        /// </summary>
        public decimal BevelBottomHeight
        {
            get { return this.decBevelBottomHeight; }
            set
            {
                this.decBevelBottomHeight = value;
                if (this.decBevelBottomHeight < 0m) this.decBevelBottomHeight = 0m;
                if (this.decBevelBottomHeight > 2147483647m) this.decBevelBottomHeight = 2147483647m;
                this.bHasBevelBottom = true;
            }
        }

        internal bool HasExtrusionColor;
        internal SLA.SLColorTransform clrExtrusionColor;
        /// <summary>
        /// The extrusion color, also known as the depth color. This is read-only.
        /// </summary>
        public System.Drawing.Color ExtrusionColor { get { return this.clrExtrusionColor.DisplayColor; } }

        internal decimal decExtrusionHeight;
        /// <summary>
        /// Extrusion height, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.
        /// The Microsoft Excel user interface uses the term "Depth".
        /// </summary>
        public decimal ExtrusionHeight
        {
            get { return this.decExtrusionHeight; }
            set
            {
                this.decExtrusionHeight = value;
                if (this.decExtrusionHeight < 0m) this.decExtrusionHeight = 0m;
                if (this.decExtrusionHeight > 2147483647m) this.decExtrusionHeight = 2147483647m;
            }
        }

        internal bool HasContourColor;
        internal SLA.SLColorTransform clrContourColor;
        /// <summary>
        /// The contour color. This is read-only.
        /// </summary>
        public System.Drawing.Color ContourColor { get { return this.clrContourColor.DisplayColor; } }

        internal decimal decContourWidth;
        /// <summary>
        /// Contour width, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.
        /// The Microsoft Excel user interface uses the term "Size".
        /// </summary>
        public decimal ContourWidth
        {
            get { return this.decContourWidth; }
            set
            {
                this.decContourWidth = value;
                if (this.decContourWidth < 0m) this.decContourWidth = 0m;
                if (this.decContourWidth > 2147483647m) this.decContourWidth = 2147483647m;
            }
        }

        /// <summary>
        /// The preset material used. Default is WarmMatte.
        /// </summary>
        public A.PresetMaterialTypeValues Material { get; set; }

        internal bool bHasLighting;
        /// <summary>
        /// Specifies if there's lighting.
        /// </summary>
        public bool HasLighting { get { return this.bHasLighting; } }

        internal A.LightRigValues vLighting;
        /// <summary>
        /// The type of lighting used.
        /// </summary>
        public A.LightRigValues Lighting
        {
            get { return this.vLighting; }
            set
            {
                this.vLighting = value;
                this.bHasLighting = true;
            }
        }

        internal decimal decAngle;
        /// <summary>
        /// Angle of the lighting, ranging from 0 degrees to 359.9 degrees. This is set only when <see cref="Lighting"/> is also set.
        /// </summary>
        public decimal Angle
        {
            get { return decAngle; }
            set
            {
                this.decAngle = value;
                if (this.decAngle < 0m) this.decAngle = 0m;
                if (this.decAngle >= 360m) this.decAngle = 359.9m;
            }
        }

        internal SLFormat3D(List<System.Drawing.Color> ThemeColors)
        {
            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }

            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.SetNoBevelTop();
            this.SetNoBevelBottom();
            this.SetNoExtrusion();
            this.SetNoContour();

            this.Material = A.PresetMaterialTypeValues.WarmMatte;

            this.SetNoLighting();
        }

        /// <summary>
        /// Set the top bevel.
        /// </summary>
        /// <param name="BevelPreset">The bevel type.</param>
        /// <param name="Width">Bevel width ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.</param>
        /// <param name="Height">Bevel height ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.</param>
        public void SetBevelTop(A.BevelPresetValues BevelPreset, decimal Width, decimal Height)
        {
            this.vBevelTopPreset = BevelPreset;
            this.BevelTopWidth = Width;
            this.BevelTopHeight = Height;
            this.bHasBevelTop = true;
        }

        /// <summary>
        /// Remove the top bevel.
        /// </summary>
        public void SetNoBevelTop()
        {
            this.vBevelTopPreset = A.BevelPresetValues.Circle;
            this.decBevelTopWidth = 6;
            this.decBevelTopHeight = 6;
            this.bHasBevelTop = false;
        }

        /// <summary>
        /// Set the bottom bevel.
        /// </summary>
        /// <param name="BevelPreset">The bevel type.</param>
        /// <param name="Width">Bevel width ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.</param>
        /// <param name="Height">Bevel height ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.</param>
        public void SetBevelBottom(A.BevelPresetValues BevelPreset, decimal Width, decimal Height)
        {
            this.vBevelBottomPreset = BevelPreset;
            this.BevelBottomWidth = Width;
            this.BevelBottomHeight = Height;
            this.bHasBevelBottom = true;
        }

        /// <summary>
        /// Remove the bottom bevel.
        /// </summary>
        public void SetNoBevelBottom()
        {
            this.vBevelBottomPreset = A.BevelPresetValues.Circle;
            this.decBevelBottomWidth = 6;
            this.decBevelBottomHeight = 6;
            this.bHasBevelBottom = false;
        }

        /// <summary>
        /// Remove any extrusion (or depth) settings.
        /// </summary>
        public void SetNoExtrusion()
        {
            this.clrExtrusionColor = new SLColorTransform(this.listThemeColors);
            this.HasExtrusionColor = false;
            this.decExtrusionHeight = 0;
        }

        /// <summary>
        /// Set the extrusion (or depth) color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        public void SetExtrusionColor(System.Drawing.Color Color)
        {
            if (!Color.IsEmpty)
            {
                this.clrExtrusionColor.SetColor(Color, 0);
                this.HasExtrusionColor = true;
            }
        }

        /// <summary>
        /// Set the extrusion (or depth) color.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetExtrusionColor(SLThemeColorIndexValues Color, double Tint)
        {
            this.clrExtrusionColor.SetColor(Color, Tint, 0);
            this.HasExtrusionColor = true;
        }

        /// <summary>
        /// Set the extrusion.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Height">Extrusion height, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.</param>
        public void SetExtrusion(System.Drawing.Color Color, decimal Height)
        {
            if (!Color.IsEmpty)
            {
                this.clrExtrusionColor.SetColor(Color, 0);
                this.HasExtrusionColor = true;
            }
            this.ExtrusionHeight = Height;
        }

        /// <summary>
        /// Set the extrusion.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Height">Extrusion height, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.</param>
        public void SetExtrusion(SLThemeColorIndexValues Color, double Tint, decimal Height)
        {
            this.clrExtrusionColor.SetColor(Color, Tint, 0);
            this.HasExtrusionColor = true;
            this.ExtrusionHeight = Height;
        }

        /// <summary>
        /// Remove any contour settings.
        /// </summary>
        public void SetNoContour()
        {
            this.clrContourColor = new SLColorTransform(this.listThemeColors);
            this.HasContourColor = false;
            this.decContourWidth = 0;
        }

        /// <summary>
        /// Set the contour color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        public void SetContourColor(System.Drawing.Color Color)
        {
            if (!Color.IsEmpty)
            {
                this.clrContourColor.SetColor(Color, 0);
                this.HasContourColor = true;
            }
        }

        /// <summary>
        /// Set the contour color.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetContourColor(SLThemeColorIndexValues Color, double Tint)
        {
            this.clrContourColor.SetColor(Color, Tint, 0);
            this.HasContourColor = true;
        }

        /// <summary>
        /// Set the contour.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Width">Contour width, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.</param>
        public void SetContour(System.Drawing.Color Color, decimal Width)
        {
            if (!Color.IsEmpty)
            {
                this.clrContourColor.SetColor(Color, 0);
                this.HasContourColor = true;
            }
            this.ContourWidth = Width;
        }

        /// <summary>
        /// Set the contour.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Width">Contour width, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of a point.</param>
        public void SetContour(SLThemeColorIndexValues Color, double Tint, decimal Width)
        {
            this.clrContourColor.SetColor(Color, Tint, 0);
            this.HasContourColor = true;
            this.ContourWidth = Width;
        }

        /// <summary>
        /// Remove any lighting settings.
        /// </summary>
        public void SetNoLighting()
        {
            this.vLighting = A.LightRigValues.ThreePoints;
            this.bHasLighting = false;
            this.decAngle = 0;
        }

        internal SLFormat3D Clone()
        {
            SLFormat3D format = new SLFormat3D(this.listThemeColors);
            format.bHasBevelTop = this.bHasBevelTop;
            format.vBevelTopPreset = this.vBevelTopPreset;
            format.decBevelTopWidth = this.decBevelTopWidth;
            format.decBevelTopHeight = this.decBevelTopHeight;
            format.bHasBevelBottom = this.bHasBevelBottom;
            format.vBevelBottomPreset = this.vBevelBottomPreset;
            format.decBevelBottomWidth = this.decBevelBottomWidth;
            format.decBevelBottomHeight = this.decBevelBottomHeight;
            format.HasExtrusionColor = this.HasExtrusionColor;
            format.clrExtrusionColor = this.clrExtrusionColor.Clone();
            format.decExtrusionHeight = this.decExtrusionHeight;
            format.HasContourColor = this.HasContourColor;
            format.clrContourColor = this.clrContourColor.Clone();
            format.decContourWidth = this.decContourWidth;
            format.Material = this.Material;
            format.bHasLighting = this.bHasLighting;
            format.vLighting = this.vLighting;
            format.decAngle = this.decAngle;

            return format;
        }
    }
}
