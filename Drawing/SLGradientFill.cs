using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing
{
    internal class SLGradientFill
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal bool ThrowExceptionsIfAny { get; set; }

        internal bool IsLinear = true;
        private decimal decAngle;
        /// <summary>
        /// The interpolation angle ranging from 0 degrees to 359.9 degrees. 0 degrees mean from left to right, 90 degrees mean from top to bottom, 180 degrees mean from right to left and 270 degrees mean from bottom to top. Accurate to 1/60000 of a degree.
        /// </summary>
        internal decimal Angle
        {
            get { return decAngle; }
            set
            {
                decAngle = value;
                if (decAngle < 0m) decAngle = 0m;
                if (decAngle >= 360m) decAngle = 359.9m;
            }
        }

        internal A.PathShadeValues PathType { get; set; }
        internal SLGradientDirectionValues Direction { get; set; }

        internal bool HasFlip = false;
        private A.TileFlipValues vFlip;
        internal A.TileFlipValues Flip
        {
            get { return vFlip; }
            set
            {
                this.HasFlip = true;
                this.vFlip = value;
            }
        }

        internal bool HasRotateWithShape = false;
        private bool bRotateWithShape;
        internal bool RotateWithShape
        {
            get { return bRotateWithShape; }
            set
            {
                this.HasRotateWithShape = true;
                this.bRotateWithShape = value;
            }
        }

        internal List<SLGradientStop> GradientStops { get; set; }

        internal SLGradientFill(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }

            this.ThrowExceptionsIfAny = ThrowExceptionsIfAny;

            this.IsLinear = true;
            this.Angle = 0;
            this.PathType = A.PathShadeValues.Circle;
            this.Direction = SLGradientDirectionValues.Center;
            this.vFlip = A.TileFlipValues.None;
            this.HasFlip = false;
            this.bRotateWithShape = true;
            this.HasRotateWithShape = false;
            this.GradientStops = new List<SLGradientStop>();
        }

        internal void SetLinearGradient(SLGradientPresetValues Preset, decimal Angle)
        {
            this.IsLinear = true;
            this.Angle = Angle;
            this.FillGradientStops(Preset);
        }

        internal void SetRadialGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            this.IsLinear = false;
            this.PathType = A.PathShadeValues.Circle;
            this.Direction = Direction;
            this.FillGradientStops(Preset);
        }

        internal void SetRectangularGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            this.IsLinear = false;
            this.PathType = A.PathShadeValues.Rectangle;
            this.Direction = Direction;
            this.FillGradientStops(Preset);
        }

        internal void SetPathGradient(SLGradientPresetValues Preset)
        {
            this.IsLinear = false;
            this.PathType = A.PathShadeValues.Shape;
            this.FillGradientStops(Preset);
        }

        internal void AppendGradientStop(System.Drawing.Color Color, decimal Transparency, decimal Position)
        {
            SLGradientStop gs = new SLGradientStop(this.listThemeColors);
            gs.Color.SetColor(Color, Transparency);
            gs.Position = Position;
            this.GradientStops.Add(gs);
        }

        internal void AppendGradientStop(SLThemeColorIndexValues Color, double Tint, decimal Transparency, decimal Position)
        {
            SLGradientStop gs = new SLGradientStop(this.listThemeColors);
            gs.Color.SetColor(Color, Tint, Transparency);
            gs.Position = Position;
            this.GradientStops.Add(gs);
        }

        internal void ClearGradientStops()
        {
            this.GradientStops.Clear();
        }

        internal void FillGradientStops(SLGradientPresetValues PresetType)
        {
            this.GradientStops = new List<SLGradientStop>();
            switch (PresetType)
            {
                case SLGradientPresetValues.EarlySunset:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "000082", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "66008F", 30, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "BA0066", 64.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FF0000", 89.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FF8200", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.LateSunset:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "000000", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "000040", 20, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "400040", 50, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "8F0040", 75, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "F27300", 89.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFBF00", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Nightfall:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "000000", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "0A128C", 39.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "181CC7", 70, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "7005D4", 88, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "8C3D91", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Daybreak:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "5E9EFF", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "85C2FF", 39.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "C4D6EB", 70, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFEBFA", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Horizon:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "DCEBF5", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "83A7C3", 8, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "768FB9", 13, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "83A7C3", 21.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFFFFF", 52, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "9C6563", 56, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "80302D", 58, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "C0524E", 71.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "EBDAD4", 94, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "55261C", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Desert:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FC9FCB", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "F8B049", 13, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "F8B049", 21.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FEE7F2", 63, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "F952A0", 67, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "C50849", 69, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "B43E85", 82.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "F8B049", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Ocean:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "03D4A8", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "21D6E0", 25, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "0087E6", 75, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "005CBF", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.CalmWater:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "CCCCFF", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "99CCFF", 17.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "9966FF", 36, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "CC99FF", 61, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "99CCFF", 82.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "CCCCFF", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Fire:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFF200", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FF7A00", 45, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FF0300", 70, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "4D0808", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Fog:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "8488C4", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "D4DEFF", 53, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "D4DEFF", 83, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "96AB94", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Moss:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "DDEBCF", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "9CB86E", 50, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "156B13", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Peacock:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "3399FF", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "00CCCC", 16, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "9999FF", 47, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "2E6792", 60.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "3333CC", 71.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "1170FF", 81, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "006699", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Wheat:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FBEAC7", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FEE7F2", 17.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FAC77D", 36, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FBA97D", 61, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FBD49C", 82.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FEE7F2", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Parchment:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFEFD1", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "F0EBD5", 64.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "D1C39F", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Mahogany:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "D6B19C", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "D49E6C", 30, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "A65528", 70, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "663012", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Rainbow:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "A603AB", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "0819FB", 21.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "1A8D48", 35.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFFF00", 52, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "EE3F17", 73, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "E81766", 88, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "A603AB", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Rainbow2:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FF3399", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FF6633", 25, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFFF00", 50, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "01A78F", 75, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "3366FF", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Gold:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "E6DCAC", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "E6D78A", 12, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "C7AC4C", 30, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "E6D78A", 45, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "C7AC4C", 77, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "E6DCAC", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Gold2:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FBE4AE", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "BD922A", 13, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "BD922A", 21.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FBE4AE", 63, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "BD922A", 67, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "835E17", 69, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "A28949", 82.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FAE3B7", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Brass:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "825600", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFA800", 13, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "825600", 28, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFA800", 42.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "825600", 58, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFA800", 72, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "825600", 87, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFA800", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Chrome:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFFFFF", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "1F1F1F", 16, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFFFFF", 17.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "636363", 42, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "CFCFCF", 53, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "CFCFCF", 66, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "1F1F1F", 75.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFFFFF", 78.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "7F7F7F", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Chrome2:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "CBCBCB", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "5F5F5F", 13, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "5F5F5F", 21.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFFFFF", 63, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "B2B2B2", 67, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "292929", 69, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "777777", 82.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "EAEAEA", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Silver:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "FFFFFF", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "E6E6E6", 7.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "7D8496", 32.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "E6E6E6", 47, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "7D8496", 85.001m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "E6E6E6", 100, this.ThrowExceptionsIfAny));
                    break;
                case SLGradientPresetValues.Sapphire:
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "000082", 0, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "0047FF", 13, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "000082", 28, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "0047FF", 42.999m, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "000082", 58, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "0047FF", 72, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "000082", 87, this.ThrowExceptionsIfAny));
                    this.GradientStops.Add(new SLGradientStop(this.listThemeColors, "0047FF", 100, this.ThrowExceptionsIfAny));
                    break;
            }
        }

        internal A.GradientFill ToGradientFill()
        {
            A.GradientFill gf = new A.GradientFill();

            A.GradientStopList gsl = new A.GradientStopList();
            for (int i = 0; i < this.GradientStops.Count; ++i)
            {
                gsl.Append(this.GradientStops[i].ToGradientStop());
            }
            gf.Append(gsl);

            if (this.IsLinear)
            {
                A.LinearGradientFill lgf = new A.LinearGradientFill();
                lgf.Angle = Convert.ToInt32(this.Angle * SLConstants.DegreeToAngleRepresentation);
                lgf.Scaled = false;
                gf.Append(lgf);
                gf.Append(new A.TileRectangle());
            }
            else
            {
                if (this.PathType == A.PathShadeValues.Shape)
                {
                    A.PathGradientFill pgf = new A.PathGradientFill();
                    pgf.Path = this.PathType;
                    pgf.FillToRectangle = new A.FillToRectangle()
                    {
                        Left = 50000,
                        Top = 50000,
                        Right = 50000,
                        Bottom = 50000
                    };
                    gf.Append(pgf);
                    gf.Append(new A.TileRectangle());
                }
                else
                {
                    A.PathGradientFill pgf = new A.PathGradientFill();
                    pgf.Path = this.PathType;
                    switch (this.Direction)
                    {
                        case SLGradientDirectionValues.CenterToTopLeftCorner:
                            pgf.FillToRectangle = new A.FillToRectangle() { Left = 100000, Top = 100000 };
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle() { Right = -100000, Bottom = -100000 });
                            break;
                        case SLGradientDirectionValues.CenterToTopRightCorner:
                            pgf.FillToRectangle = new A.FillToRectangle() { Top = 100000, Right = 100000 };
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle() { Left = -100000, Bottom = -100000 });
                            break;
                        case SLGradientDirectionValues.Center:
                            pgf.FillToRectangle = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle());
                            break;
                        case SLGradientDirectionValues.CenterToBottomLeftCorner:
                            pgf.FillToRectangle = new A.FillToRectangle() { Left = 100000, Bottom = 100000 };
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle() { Top = -100000, Right = -100000 });
                            break;
                        case SLGradientDirectionValues.CenterToBottomRightCorner:
                            pgf.FillToRectangle = new A.FillToRectangle() { Right = 100000, Bottom = 100000 };
                            gf.Append(pgf);
                            gf.Append(new A.TileRectangle() { Left = -100000, Top = -100000 });
                            break;
                    }
                }
            }

            if (this.HasFlip) gf.Flip = this.Flip;
            if (this.HasRotateWithShape) gf.RotateWithShape = this.RotateWithShape;

            return gf;
        }

        internal SLGradientFill Clone()
        {
            SLGradientFill gf = new SLGradientFill(this.listThemeColors, this.ThrowExceptionsIfAny);
            gf.IsLinear = this.IsLinear;
            gf.decAngle = this.decAngle;
            gf.PathType = this.PathType;
            gf.Direction = this.Direction;
            gf.HasFlip = this.HasFlip;
            gf.vFlip = this.vFlip;
            gf.HasRotateWithShape = this.HasRotateWithShape;
            gf.bRotateWithShape = this.bRotateWithShape;
            for (int i = 0; i < this.GradientStops.Count; ++i)
            {
                gf.GradientStops.Add(this.GradientStops[i].Clone());
            }

            return gf;
        }
    }
}
