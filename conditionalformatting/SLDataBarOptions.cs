using System;
using System.Collections.Generic;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    /// <summary>
    /// Conditional formatting options for data bars.
    /// </summary>
    public class SLDataBarOptions
    {
        internal bool Is2010;

        internal SLConditionalFormatAutoMinMaxValues vMinimumType;
        /// <summary>
        /// The conditional format type for the minimum value. If "Automatic" is used, Excel 2010 specific data bars will be used.
        /// </summary>
        public SLConditionalFormatAutoMinMaxValues MinimumType
        {
            get { return vMinimumType; }
            set
            {
                vMinimumType = value;
                if (vMinimumType == SLConditionalFormatAutoMinMaxValues.Automatic) this.Is2010 = true;
            }
        }

        /// <summary>
        /// The minimum value.
        /// </summary>
        public string MinimumValue { get; set; }

        internal SLConditionalFormatAutoMinMaxValues vMaximumType;
        /// <summary>
        /// The conditional format type for the maximum value. If "Automatic" is used, Excel 2010 specific data bars will be used.
        /// </summary>
        public SLConditionalFormatAutoMinMaxValues MaximumType
        {
            get { return vMaximumType; }
            set
            {
                vMaximumType = value;
                if (vMaximumType == SLConditionalFormatAutoMinMaxValues.Automatic) this.Is2010 = true;
            }
        }

        /// <summary>
        /// The maximum value.
        /// </summary>
        public string MaximumValue { get; set; }

        /// <summary>
        /// The fill color.
        /// </summary>
        public SLColor FillColor { get; set; }

        /// <summary>
        /// The border color.
        /// </summary>
        public SLColor BorderColor { get; set; }

        /// <summary>
        /// The fill color for negative values.
        /// </summary>
        public SLColor NegativeFillColor { get; set; }

        /// <summary>
        /// The border color for negative values.
        /// </summary>
        public SLColor NegativeBorderColor { get; set; }

        /// <summary>
        /// The axis color.
        /// </summary>
        public SLColor AxisColor { get; set; }

        /// <summary>
        /// The minimum length of the data bar as a percentage of the cell width. The default value is 10.
        /// </summary>
        public uint MinLength { get; set; }

        /// <summary>
        /// The maximum length of the data bar as a percentage of the cell width. The default value is 90. It is recommended to keep this to a maximum (haha) of 100.
        /// </summary>
        public uint MaxLength { get; set; }
        
        /// <summary>
        /// Specifies if only the data bar is shown. Set to false to show both data bar and value.
        /// </summary>
        public bool ShowBarOnly { get; set; }

        internal bool bBorder;
        /// <summary>
        /// Specifies if there's a border. This is an Excel 2010 specific feature.
        /// </summary>
        public bool Border
        {
            get { return bBorder; }
            set
            {
                bBorder = value;
                Is2010 = true;
            }
        }

        internal bool bGradient;
        /// <summary>
        /// Specifies if the fill color has a gradient. This is an Excel 2010 specific feature.
        /// </summary>
        public bool Gradient
        {
            get { return bGradient; }
            set
            {
                bGradient = value;
                Is2010 = true;
            }
        }

        internal X14.DataBarDirectionValues vDirection;
        /// <summary>
        /// The bar direction. This is an Excel 2010 specific feature.
        /// </summary>
        public X14.DataBarDirectionValues Direction
        {
            get { return vDirection; }
            set
            {
                vDirection = value;
                Is2010 = true;
            }
        }

        internal bool bNegativeBarColorSameAsPositive;
        /// <summary>
        /// Specifies if the fill color for negative values is the same as the positive one. This is an Excel 2010 specific feature.
        /// </summary>
        public bool NegativeBarColorSameAsPositive
        {
            get { return bNegativeBarColorSameAsPositive; }
            set
            {
                bNegativeBarColorSameAsPositive = value;
                Is2010 = true;
            }
        }

        internal bool bNegativeBarBorderColorSameAsPositive;
        /// <summary>
        /// Specifies if the border color for negative values is the same as the positive one. This is an Excel 2010 specific feature.
        /// </summary>
        public bool NegativeBarBorderColorSameAsPositive
        {
            get { return bNegativeBarBorderColorSameAsPositive; }
            set
            {
                bNegativeBarBorderColorSameAsPositive = value;
                Is2010 = true;
            }
        }

        internal X14.DataBarAxisPositionValues vAxisPosition;
        /// <summary>
        /// Specifies the axis position. This is an Excel 2010 specific feature.
        /// </summary>
        public X14.DataBarAxisPositionValues AxisPosition
        {
            get { return vAxisPosition; }
            set
            {
                vAxisPosition = value;
                Is2010 = true;
            }
        }

        /// <summary>
        /// Initializes an instance of SLDataBarOptions.
        /// </summary>
        public SLDataBarOptions()
        {
            this.InitialiseDataBarOptions(SLConditionalFormatDataBarValues.Blue, true);
        }

        /// <summary>
        /// Initializes an instance of SLDataBarOptions.
        /// </summary>
        /// <param name="DataBar">Built-in data bar type.</param>
        public SLDataBarOptions(SLConditionalFormatDataBarValues DataBar)
        {
            this.InitialiseDataBarOptions(DataBar, true);
        }

        /// <summary>
        /// Initializes an instance of SLDataBarOptions.
        /// </summary>
        /// <param name="Is2010Default">True if Excel 2010 specific data bar is to be used. False otherwise.</param>
        public SLDataBarOptions(bool Is2010Default)
        {
            this.InitialiseDataBarOptions(SLConditionalFormatDataBarValues.Blue, Is2010Default);
        }

        /// <summary>
        /// Initializes an instance of SLDataBarOptions.
        /// </summary>
        /// <param name="DataBar">Built-in data bar type.</param>
        /// <param name="Is2010Default">True if Excel 2010 specific data bar is to be used. False otherwise.</param>
        public SLDataBarOptions(SLConditionalFormatDataBarValues DataBar, bool Is2010Default)
        {
            this.InitialiseDataBarOptions(DataBar, Is2010Default);
        }

        private void InitialiseDataBarOptions(SLConditionalFormatDataBarValues DataBar, bool Is2010Default)
        {
            this.Is2010 = Is2010Default;

            this.FillColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.BorderColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.NegativeFillColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.NegativeBorderColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
            this.AxisColor = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());

            switch (DataBar)
            {
                case SLConditionalFormatDataBarValues.Blue:
                    this.FillColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0x8E, 0xC6);
                    this.BorderColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0x8E, 0xC6);
                    break;
                case SLConditionalFormatDataBarValues.Green:
                    this.FillColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xC3, 0x84);
                    this.BorderColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xC3, 0x84);
                    break;
                case SLConditionalFormatDataBarValues.Red:
                    this.FillColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x55, 0x5A);
                    this.BorderColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x55, 0x5A);
                    break;
                case SLConditionalFormatDataBarValues.Orange:
                    this.FillColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xB6, 0x28);
                    this.BorderColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xB6, 0x28);
                    break;
                case SLConditionalFormatDataBarValues.LightBlue:
                    this.FillColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x00, 0x8A, 0xEF);
                    this.BorderColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x00, 0x8A, 0xEF);
                    break;
                case SLConditionalFormatDataBarValues.Purple:
                    this.FillColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD6, 0x00, 0x7B);
                    this.BorderColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xD6, 0x00, 0x7B);
                    break;
            }

            this.NegativeFillColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x00, 0x00);
            this.NegativeBorderColor.Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x00, 0x00);
            this.AxisColor.Color = System.Drawing.Color.FromArgb(0xFF, 0x00, 0x00, 0x00);

            if (Is2010Default)
            {
                this.vMinimumType = SLConditionalFormatAutoMinMaxValues.Automatic;
                this.MinimumValue = string.Empty;
                this.vMaximumType = SLConditionalFormatAutoMinMaxValues.Automatic;
                this.MaximumValue = string.Empty;
                this.MinLength = 0;
                this.MaxLength = 100;
            }
            else
            {
                this.vMinimumType = SLConditionalFormatAutoMinMaxValues.Value;
                this.MinimumValue = string.Empty;
                this.vMaximumType = SLConditionalFormatAutoMinMaxValues.Value;
                this.MaximumValue = string.Empty;
                this.MinLength = 10;
                this.MaxLength = 90;
            }

            this.ShowBarOnly = false;
            this.bBorder = false;
            this.bGradient = false;
            this.vDirection = X14.DataBarDirectionValues.Context;
            this.bNegativeBarColorSameAsPositive = false;
            this.bNegativeBarBorderColorSameAsPositive = true;
            this.vAxisPosition = X14.DataBarAxisPositionValues.Automatic;
        }
    }
}
