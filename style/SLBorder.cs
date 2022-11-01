using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for specifying cell borders. This simulates the DocumentFormat.OpenXml.Spreadsheet.Border class.
    /// </summary>
    public class SLBorder
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        internal bool HasLeftBorder;
        internal SLBorderProperties bpLeftBorder;
        /// <summary>
        /// Encapsulates properties and methods for specifying the left border.
        /// </summary>
        public SLBorderProperties LeftBorder
        {
            get { return bpLeftBorder; }
            set
            {
                bpLeftBorder = value;
                HasLeftBorder = true;
            }
        }

        internal bool HasRightBorder;
        internal SLBorderProperties bpRightBorder;
        /// <summary>
        /// Encapsulates properties and methods for specifying the right border.
        /// </summary>
        public SLBorderProperties RightBorder
        {
            get { return bpRightBorder; }
            set
            {
                bpRightBorder = value;
                HasRightBorder = true;
            }
        }

        internal bool HasTopBorder;
        internal SLBorderProperties bpTopBorder;
        /// <summary>
        /// Encapsulates properties and methods for specifying the top border.
        /// </summary>
        public SLBorderProperties TopBorder
        {
            get { return bpTopBorder; }
            set
            {
                bpTopBorder = value;
                HasTopBorder = true;
            }
        }

        internal bool HasBottomBorder;
        internal SLBorderProperties bpBottomBorder;
        /// <summary>
        /// Encapsulates properties and methods for specifying the bottom border.
        /// </summary>
        public SLBorderProperties BottomBorder
        {
            get { return bpBottomBorder; }
            set
            {
                bpBottomBorder = value;
                HasBottomBorder = true;
            }
        }

        internal bool HasDiagonalBorder;
        internal SLBorderProperties bpDiagonalBorder;
        /// <summary>
        /// Encapsulates properties and methods for specifying the diagonal border.
        /// </summary>
        public SLBorderProperties DiagonalBorder
        {
            get { return bpDiagonalBorder; }
            set
            {
                bpDiagonalBorder = value;
                HasDiagonalBorder = true;
            }
        }

        internal bool HasVerticalBorder;
        internal SLBorderProperties bpVerticalBorder;
        /// <summary>
        /// Encapsulates properties and methods for specifying the vertical border.
        /// </summary>
        public SLBorderProperties VerticalBorder
        {
            get { return bpVerticalBorder; }
            set
            {
                bpVerticalBorder = value;
                HasVerticalBorder = true;
            }
        }

        internal bool HasHorizontalBorder;
        internal SLBorderProperties bpHorizontalBorder;
        /// <summary>
        /// Encapsulates properties and methods for specifying the horizontal border.
        /// </summary>
        public SLBorderProperties HorizontalBorder
        {
            get { return bpHorizontalBorder; }
            set
            {
                bpHorizontalBorder = value;
                HasHorizontalBorder = true;
            }
        }

        /// <summary>
        /// Specifies if there's a diagonal line from the bottom left corner of the cell to the top right corner of the cell.
        /// </summary>
        public bool? DiagonalUp { get; set; }

        /// <summary>
        /// Specifies if there's a diagonal line from the top left corner of the cell to the bottom right corner of the cell.
        /// </summary>
        public bool? DiagonalDown { get; set; }

        /// <summary>
        /// Specifies if the left, right, top and bottom borders should be applied to the outside borders of a cell range.
        /// </summary>
        public bool? Outline { get; set; }

        /// <summary>
        /// Initializes an instance of SLBorder. It is recommended to use CreateBorder() of the SLDocument class.
        /// </summary>
        public SLBorder()
        {
            this.Initialize(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
        }

        internal SLBorder(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
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
        }

        private void SetAllNull()
        {
            RemoveLeftBorder();
            RemoveRightBorder();
            RemoveTopBorder();
            RemoveBottomBorder();
            RemoveDiagonalBorder();
            RemoveVerticalBorder();
            RemoveHorizontalBorder();

            this.DiagonalUp = null;
            this.DiagonalDown = null;
            this.Outline = null;
        }

        /// <summary>
        /// Remove any existing left border.
        /// </summary>
        public void RemoveLeftBorder()
        {
            this.bpLeftBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
            HasLeftBorder = false;
        }

        /// <summary>
        /// Remove any existing right border.
        /// </summary>
        public void RemoveRightBorder()
        {
            this.bpRightBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
            HasRightBorder = false;
        }

        /// <summary>
        /// Remove any existing top border.
        /// </summary>
        public void RemoveTopBorder()
        {
            this.bpTopBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
            HasTopBorder = false;
        }

        /// <summary>
        /// Remove any existing bottom border.
        /// </summary>
        public void RemoveBottomBorder()
        {
            this.bpBottomBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
            HasBottomBorder = false;
        }

        /// <summary>
        /// Remove any existing diagonal border.
        /// </summary>
        public void RemoveDiagonalBorder()
        {
            this.bpDiagonalBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
            HasDiagonalBorder = false;
        }

        /// <summary>
        /// Remove any existing vertical border.
        /// </summary>
        public void RemoveVerticalBorder()
        {
            this.bpVerticalBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
            HasVerticalBorder = false;
        }

        /// <summary>
        /// Remove any existing horizontal border.
        /// </summary>
        public void RemoveHorizontalBorder()
        {
            this.bpHorizontalBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
            HasHorizontalBorder = false;
        }

        /// <summary>
        /// Remove all borders.
        /// </summary>
        public void RemoveAllBorders()
        {
            this.RemoveLeftBorder();
            this.RemoveRightBorder();
            this.RemoveTopBorder();
            this.RemoveBottomBorder();
            this.RemoveDiagonalBorder();
            this.RemoveVerticalBorder();
            this.RemoveHorizontalBorder();
        }

        /// <summary>
        /// Set the left border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetLeftBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.LeftBorder.BorderStyle = BorderStyle;
            this.LeftBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the left border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetLeftBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.LeftBorder.BorderStyle = BorderStyle;
            this.LeftBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the left border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetLeftBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.LeftBorder.BorderStyle = BorderStyle;
            this.LeftBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the right border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetRightBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.RightBorder.BorderStyle = BorderStyle;
            this.RightBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the right border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetRightBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.RightBorder.BorderStyle = BorderStyle;
            this.RightBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the right border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetRightBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.RightBorder.BorderStyle = BorderStyle;
            this.RightBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the top border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetTopBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.TopBorder.BorderStyle = BorderStyle;
            this.TopBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the top border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetTopBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.TopBorder.BorderStyle = BorderStyle;
            this.TopBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the top border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetTopBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.TopBorder.BorderStyle = BorderStyle;
            this.TopBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the bottom border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetBottomBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.BottomBorder.BorderStyle = BorderStyle;
            this.BottomBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the bottom border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetBottomBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.BottomBorder.BorderStyle = BorderStyle;
            this.BottomBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the bottom border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetBottomBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.BottomBorder.BorderStyle = BorderStyle;
            this.BottomBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the diagonal border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetDiagonalBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.DiagonalBorder.BorderStyle = BorderStyle;
            this.DiagonalBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the diagonal border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetDiagonalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.DiagonalBorder.BorderStyle = BorderStyle;
            this.DiagonalBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the diagonal border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetDiagonalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.DiagonalBorder.BorderStyle = BorderStyle;
            this.DiagonalBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the vertical border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetVerticalBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.VerticalBorder.BorderStyle = BorderStyle;
            this.VerticalBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the vertical border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetVerticalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.VerticalBorder.BorderStyle = BorderStyle;
            this.VerticalBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the vertical border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetVerticalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.VerticalBorder.BorderStyle = BorderStyle;
            this.VerticalBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        /// Set the horizontal border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetHorizontalBorder(BorderStyleValues BorderStyle, System.Drawing.Color BorderColor)
        {
            this.HorizontalBorder.BorderStyle = BorderStyle;
            this.HorizontalBorder.Color = BorderColor;
        }

        /// <summary>
        /// Set the horizontal border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetHorizontalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            this.HorizontalBorder.BorderStyle = BorderStyle;
            this.HorizontalBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        /// Set the horizontal border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetHorizontalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            this.HorizontalBorder.BorderStyle = BorderStyle;
            this.HorizontalBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        internal void Sync()
        {
            HasLeftBorder = LeftBorder.HasColor || LeftBorder.HasBorderStyle;
            HasRightBorder = RightBorder.HasColor || RightBorder.HasBorderStyle;
            HasTopBorder = TopBorder.HasColor || TopBorder.HasBorderStyle;
            HasBottomBorder = BottomBorder.HasColor || BottomBorder.HasBorderStyle;
            HasDiagonalBorder = DiagonalBorder.HasColor || DiagonalBorder.HasBorderStyle;
            HasVerticalBorder = VerticalBorder.HasColor || VerticalBorder.HasBorderStyle;
            HasHorizontalBorder = HorizontalBorder.HasColor || HorizontalBorder.HasBorderStyle;
        }

        /// <summary>
        /// Form SLBorder from DocumentFormat.OpenXml.Spreadsheet.Border class.
        /// </summary>
        /// <param name="border">The source DocumentFormat.OpenXml.Spreadsheet.Border class.</param>
        public void FromBorder(Border border)
        {
            this.SetAllNull();

            if (border.LeftBorder != null)
            {
                HasLeftBorder = true;
                this.bpLeftBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
                this.bpLeftBorder.FromBorderPropertiesType(border.LeftBorder);
            }

            if (border.RightBorder != null)
            {
                HasRightBorder = true;
                this.bpRightBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
                this.bpRightBorder.FromBorderPropertiesType(border.RightBorder);
            }

            if (border.TopBorder != null)
            {
                HasTopBorder = true;
                this.bpTopBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
                this.bpTopBorder.FromBorderPropertiesType(border.TopBorder);
            }

            if (border.BottomBorder != null)
            {
                HasBottomBorder = true;
                this.bpBottomBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
                this.bpBottomBorder.FromBorderPropertiesType(border.BottomBorder);
            }

            if (border.DiagonalBorder != null)
            {
                HasDiagonalBorder = true;
                this.bpDiagonalBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
                this.bpDiagonalBorder.FromBorderPropertiesType(border.DiagonalBorder);
            }

            if (border.VerticalBorder != null)
            {
                HasVerticalBorder = true;
                this.bpVerticalBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
                this.bpVerticalBorder.FromBorderPropertiesType(border.VerticalBorder);
            }

            if (border.HorizontalBorder != null)
            {
                HasHorizontalBorder = true;
                this.bpHorizontalBorder = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
                this.bpHorizontalBorder.FromBorderPropertiesType(border.HorizontalBorder);
            }

            if (border.DiagonalUp != null) this.DiagonalUp = border.DiagonalUp.Value;
            else this.DiagonalUp = null;

            if (border.DiagonalDown != null) this.DiagonalDown = border.DiagonalDown.Value;
            else this.DiagonalDown = null;

            if (border.Outline != null) this.Outline = border.Outline.Value;
            else this.Outline = null;

            Sync();
        }

        /// <summary>
        /// Form a DocumentFormat.OpenXml.Spreadsheet.Border class from SLBorder.
        /// </summary>
        /// <returns>A DocumentFormat.OpenXml.Spreadsheet.Border with the properties of this SLBorder class.</returns>
        public Border ToBorder()
        {
            Sync();

            Border border = new Border();
            // by "default" always have left, right, top, bottom and diagonal borders, even if empty?
            border.LeftBorder = this.LeftBorder.ToLeftBorder();
            border.RightBorder = this.RightBorder.ToRightBorder();
            border.TopBorder = this.TopBorder.ToTopBorder();
            border.BottomBorder = this.BottomBorder.ToBottomBorder();
            border.DiagonalBorder = this.DiagonalBorder.ToDiagonalBorder();
            if (HasVerticalBorder) border.VerticalBorder = this.VerticalBorder.ToVerticalBorder();
            if (HasHorizontalBorder) border.HorizontalBorder = this.HorizontalBorder.ToHorizontalBorder();
            if (this.DiagonalUp != null) border.DiagonalUp = this.DiagonalUp.Value;
            if (this.DiagonalDown != null) border.DiagonalDown = this.DiagonalDown.Value;
            // default is true. So set property only if false
            // This reduces tag attributes
            if (this.Outline != null && !this.Outline.Value) border.Outline = false;

            return border;
        }

        internal void FromHash(string Hash)
        {
            Border b = new Border();

            string[] saElementAttribute = Hash.Split(new string[] { SLConstants.XmlBorderElementAttributeSeparator }, StringSplitOptions.None);

            if (saElementAttribute.Length >= 2)
            {
                b.InnerXml = saElementAttribute[0];
                string[] sa = saElementAttribute[1].Split(new string[] { SLConstants.XmlBorderAttributeSeparator }, StringSplitOptions.None);
                if (sa.Length >= 3)
                {
                    if (!sa[0].Equals("null")) b.DiagonalUp = bool.Parse(sa[0]);

                    if (!sa[1].Equals("null")) b.DiagonalDown = bool.Parse(sa[1]);

                    if (!sa[2].Equals("null")) b.Outline = bool.Parse(sa[2]);
                }
            }

            this.FromBorder(b);
        }

        internal string ToHash()
        {
            Border b = this.ToBorder();
            string sXml = SLTool.RemoveNamespaceDeclaration(b.InnerXml);

            StringBuilder sb = new StringBuilder();

            sb.AppendFormat("{0}{1}", sXml, SLConstants.XmlBorderElementAttributeSeparator);

            if (b.DiagonalUp != null) sb.AppendFormat("{0}{1}", b.DiagonalUp.Value, SLConstants.XmlBorderAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlBorderAttributeSeparator);

            if (b.DiagonalDown != null) sb.AppendFormat("{0}{1}", b.DiagonalDown.Value, SLConstants.XmlBorderAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlBorderAttributeSeparator);

            if (b.Outline != null) sb.AppendFormat("{0}{1}", b.Outline.Value, SLConstants.XmlBorderAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlBorderAttributeSeparator);

            return sb.ToString();
        }

        internal SLBorder Clone()
        {
            SLBorder b = new SLBorder(this.listThemeColors, this.listIndexedColors);
            b.HasLeftBorder = this.HasLeftBorder;
            b.bpLeftBorder = this.bpLeftBorder.Clone();
            b.HasRightBorder = this.HasRightBorder;
            b.bpRightBorder = this.bpRightBorder.Clone();
            b.HasTopBorder = this.HasTopBorder;
            b.bpTopBorder = this.bpTopBorder.Clone();
            b.HasBottomBorder = this.HasBottomBorder;
            b.bpBottomBorder = this.bpBottomBorder.Clone();
            b.HasDiagonalBorder = this.HasDiagonalBorder;
            b.bpDiagonalBorder = this.bpDiagonalBorder.Clone();
            b.HasVerticalBorder = this.HasVerticalBorder;
            b.bpVerticalBorder = this.bpVerticalBorder.Clone();
            b.HasHorizontalBorder = this.HasHorizontalBorder;
            b.bpHorizontalBorder = this.bpHorizontalBorder.Clone();
            b.DiagonalUp = this.DiagonalUp;
            b.DiagonalDown = this.DiagonalDown;
            b.Outline = this.Outline;

            return b;
        }
    }

    /// <summary>
    /// Encapsulates properties and methods of border properties. This simulates the (abstract) DocumentFormat.OpenXml.Spreadsheet.BorderPropertiesType class.
    /// </summary>
    public class SLBorderProperties
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal List<System.Drawing.Color> listIndexedColors;

        internal bool HasColor;
        internal SLColor clrReal;
        /// <summary>
        /// The border color.
        /// </summary>
        public System.Drawing.Color Color
        {
            get { return clrReal.Color; }
            set
            {
                clrReal.Color = value;
                HasColor = (clrReal.Color.IsEmpty) ? false : true;
            }
        }

        internal bool HasBorderStyle;
        private BorderStyleValues vBorderStyle;
        /// <summary>
        /// The border style. Default is none.
        /// </summary>
        public BorderStyleValues BorderStyle
        {
            get { return vBorderStyle; }
            set
            {
                vBorderStyle = value;
                HasBorderStyle = vBorderStyle != BorderStyleValues.None ? true : false;
            }
        }

        internal SLBorderProperties(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors)
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

            RemoveColor();
            RemoveBorderStyle();
        }

        /// <summary>
        /// Set the color of the border with one of the theme colors.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        public void SetBorderThemeColor(SLThemeColorIndexValues ThemeColorIndex)
        {
            this.clrReal.SetThemeColor(ThemeColorIndex);
            HasColor = (clrReal.Color.IsEmpty) ? false : true;
        }

        /// <summary>
        /// Set the color of the border with one of the theme colors, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetBorderThemeColor(SLThemeColorIndexValues ThemeColorIndex, double Tint)
        {
            this.clrReal.SetThemeColor(ThemeColorIndex, Tint);
            HasColor = (clrReal.Color.IsEmpty) ? false : true;
        }

        /// <summary>
        /// Remove any existing color.
        /// </summary>
        public void RemoveColor()
        {
            this.clrReal = new SLColor(this.listThemeColors, this.listIndexedColors);
            HasColor = false;
        }

        /// <summary>
        /// Remove any existing border style.
        /// </summary>
        public void RemoveBorderStyle()
        {
            this.vBorderStyle = BorderStyleValues.None;
            HasBorderStyle = false;
        }

        internal void FromBorderPropertiesType(LeftBorder border)
        {
            if (border.Color != null)
            {
                this.clrReal = new SLColor(this.listThemeColors, this.listIndexedColors);
                this.clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !this.clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) this.BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(RightBorder border)
        {
            if (border.Color != null)
            {
                this.clrReal = new SLColor(this.listThemeColors, this.listIndexedColors);
                this.clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !this.clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) this.BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(TopBorder border)
        {
            if (border.Color != null)
            {
                this.clrReal = new SLColor(this.listThemeColors, this.listIndexedColors);
                this.clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !this.clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) this.BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(BottomBorder border)
        {
            if (border.Color != null)
            {
                this.clrReal = new SLColor(this.listThemeColors, this.listIndexedColors);
                this.clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !this.clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) this.BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(DiagonalBorder border)
        {
            if (border.Color != null)
            {
                this.clrReal = new SLColor(this.listThemeColors, this.listIndexedColors);
                this.clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !this.clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) this.BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(VerticalBorder border)
        {
            if (border.Color != null)
            {
                this.clrReal = new SLColor(this.listThemeColors, this.listIndexedColors);
                this.clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !this.clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) this.BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(HorizontalBorder border)
        {
            if (border.Color != null)
            {
                this.clrReal = new SLColor(this.listThemeColors, this.listIndexedColors);
                this.clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !this.clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) this.BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal LeftBorder ToLeftBorder()
        {
            LeftBorder border = new LeftBorder();
            if (HasColor) border.Color = this.clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = this.BorderStyle;

            return border;
        }

        internal RightBorder ToRightBorder()
        {
            RightBorder border = new RightBorder();
            if (HasColor) border.Color = this.clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = this.BorderStyle;

            return border;
        }

        internal TopBorder ToTopBorder()
        {
            TopBorder border = new TopBorder();
            if (HasColor) border.Color = this.clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = this.BorderStyle;

            return border;
        }

        internal BottomBorder ToBottomBorder()
        {
            BottomBorder border = new BottomBorder();
            if (HasColor) border.Color = this.clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = this.BorderStyle;

            return border;
        }

        internal DiagonalBorder ToDiagonalBorder()
        {
            DiagonalBorder border = new DiagonalBorder();
            if (HasColor) border.Color = this.clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = this.BorderStyle;

            return border;
        }

        internal VerticalBorder ToVerticalBorder()
        {
            VerticalBorder border = new VerticalBorder();
            if (HasColor) border.Color = this.clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = this.BorderStyle;

            return border;
        }

        internal HorizontalBorder ToHorizontalBorder()
        {
            HorizontalBorder border = new HorizontalBorder();
            if (HasColor) border.Color = this.clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = this.BorderStyle;

            return border;
        }

        internal void FromHash(string Hash)
        {
            // Just use the left border. Make sure it's consistent with the ToHash() function.
            LeftBorder lb = new LeftBorder();

            string[] saElementAttribute = Hash.Split(new string[] { SLConstants.XmlBorderPropertiesElementAttributeSeparator }, StringSplitOptions.None);

            if (saElementAttribute.Length >= 2)
            {
                lb.InnerXml = saElementAttribute[0];
                string[] sa = saElementAttribute[1].Split(new string[] { SLConstants.XmlBorderPropertiesAttributeSeparator }, StringSplitOptions.None);
                if (sa.Length >= 1)
                {
                    if (!sa[0].Equals("null")) lb.Style = (BorderStyleValues)Enum.Parse(typeof(BorderStyleValues), sa[0]);
                }
            }

            this.FromBorderPropertiesType(lb);
        }

        internal string ToHash()
        {
            // Just use the left border. Make sure it's consistent with the FromHash() function.
            LeftBorder lb = this.ToLeftBorder();
            string sXml = SLTool.RemoveNamespaceDeclaration(lb.InnerXml);

            StringBuilder sb = new StringBuilder();

            sb.AppendFormat("{0}{1}", sXml, SLConstants.XmlBorderPropertiesElementAttributeSeparator);

            if (lb.Style != null) sb.AppendFormat("{0}{1}", lb.Style.Value, SLConstants.XmlBorderPropertiesAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlBorderPropertiesAttributeSeparator);

            return sb.ToString();
        }

        internal static string WriteToXmlTag(string BorderTag, SLBorderProperties bp)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("<x:{0}", BorderTag);
            if (bp.HasBorderStyle)
            {
                sb.AppendFormat(" style=\"{0}\"", bp.GetBorderStyleAttribute(bp.BorderStyle));
            }

            if (bp.HasColor)
            {
                sb.Append("><x:color");
                if (bp.clrReal.Auto != null) sb.AppendFormat(" auto=\"{0}\"", bp.clrReal.Auto.Value ? "1" : "0");
                if (bp.clrReal.Indexed != null) sb.AppendFormat(" indexed=\"{0}\"", bp.clrReal.Indexed.Value);
                if (bp.clrReal.Rgb != null) sb.AppendFormat(" rgb=\"{0}\"", bp.clrReal.Rgb);
                if (bp.clrReal.Theme != null) sb.AppendFormat(" theme=\"{0}\"", bp.clrReal.Theme.Value);
                if (bp.clrReal.Tint != null) sb.AppendFormat(" tint=\"{0}\"", bp.clrReal.Tint.Value);
                sb.AppendFormat(" /></x:{0}>", BorderTag);
            }
            else
            {
                sb.Append(" />");
            }

            return sb.ToString();
        }

        internal string GetBorderStyleAttribute(BorderStyleValues bsv)
        {
            string result = "none";
            switch (bsv)
            {
                case BorderStyleValues.DashDot:
                    result = "dashDot";
                    break;
                case BorderStyleValues.DashDotDot:
                    result = "dashDotDot";
                    break;
                case BorderStyleValues.Dashed:
                    result = "dashed";
                    break;
                case BorderStyleValues.Dotted:
                    result = "dotted";
                    break;
                case BorderStyleValues.Double:
                    result = "double";
                    break;
                case BorderStyleValues.Hair:
                    result = "hair";
                    break;
                case BorderStyleValues.Medium:
                    result = "medium";
                    break;
                case BorderStyleValues.MediumDashDot:
                    result = "mediumDashDot";
                    break;
                case BorderStyleValues.MediumDashDotDot:
                    result = "mediumDashDotDot";
                    break;
                case BorderStyleValues.MediumDashed:
                    result = "mediumDashed";
                    break;
                case BorderStyleValues.None:
                    result = "none";
                    break;
                case BorderStyleValues.SlantDashDot:
                    result = "slantDashDot";
                    break;
                case BorderStyleValues.Thick:
                    result = "thick";
                    break;
                case BorderStyleValues.Thin:
                    result = "thin";
                    break;
            }

            return result;
        }

        internal SLBorderProperties Clone()
        {
            SLBorderProperties bp = new SLBorderProperties(this.listThemeColors, this.listIndexedColors);
            bp.HasColor = this.HasColor;
            bp.clrReal = this.clrReal.Clone();
            bp.HasBorderStyle = this.HasBorderStyle;
            bp.vBorderStyle = this.vBorderStyle;

            return bp;
        }
    }
}
