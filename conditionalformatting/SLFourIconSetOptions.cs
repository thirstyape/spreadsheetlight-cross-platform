using System;
using System.Collections.Generic;

namespace SpreadsheetLight
{
    /// <summary>
    /// Conditional formatting options for four icon sets.
    /// </summary>
    public class SLFourIconSetOptions
    {
        internal SLFourIconSetValues IconSetType { get; set; }

        /// <summary>
        /// Specifies if the icons in the set are reversed.
        /// </summary>
        public bool ReverseIconOrder { get; set; }
        
        /// <summary>
        /// Specifies if only the icon is shown. Set to false to show both icon and value.
        /// </summary>
        public bool ShowIconOnly { get; set; }

        internal bool IsCustomIcon;

        internal SLIconValues vIcon1;
        /// <summary>
        /// The 1st icon.
        /// </summary>
        public SLIconValues Icon1
        {
            get { return vIcon1; }
            set
            {
                if (vIcon1 != value)
                {
                    vIcon1 = value;
                    IsCustomIcon = true;
                }
            }
        }

        internal SLIconValues vIcon2;
        /// <summary>
        /// The 2nd icon.
        /// </summary>
        public SLIconValues Icon2
        {
            get { return vIcon2; }
            set
            {
                if (vIcon2 != value)
                {
                    vIcon2 = value;
                    IsCustomIcon = true;
                }
            }
        }

        internal SLIconValues vIcon3;
        /// <summary>
        /// The 3rd icon.
        /// </summary>
        public SLIconValues Icon3
        {
            get { return vIcon3; }
            set
            {
                if (vIcon3 != value)
                {
                    vIcon3 = value;
                    IsCustomIcon = true;
                }
            }
        }

        internal SLIconValues vIcon4;
        /// <summary>
        /// The 4th icon.
        /// </summary>
        public SLIconValues Icon4
        {
            get { return vIcon4; }
            set
            {
                if (vIcon4 != value)
                {
                    vIcon4 = value;
                    IsCustomIcon = true;
                }
            }
        }

        /// <summary>
        /// Specifies if values are to be greater than or equal to the 2nd range value. Set to false if values are to be strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual2 { get; set; }

        /// <summary>
        /// Specifies if values are to be greater than or equal to the 3rd range value. Set to false if values are to be strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual3 { get; set; }

        /// <summary>
        /// Specifies if values are to be greater than or equal to the 4th range value. Set to false if values are to be strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual4 { get; set; }

        /// <summary>
        /// The 2nd range value.
        /// </summary>
        public string Value2 { get; set; }

        /// <summary>
        /// The 3rd range value.
        /// </summary>
        public string Value3 { get; set; }

        /// <summary>
        /// The 4th range value.
        /// </summary>
        public string Value4 { get; set; }

        /// <summary>
        /// The conditional format type for the 2nd range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type2 { get; set; }

        /// <summary>
        /// The conditional format type for the 3rd range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type3 { get; set; }

        /// <summary>
        /// The conditional format type for the 4th range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type4 { get; set; }

        /// <summary>
        /// Initializes an instance of SLFourIconSetOptions.
        /// </summary>
        /// <param name="IconSetType">The type of icon set.</param>
        public SLFourIconSetOptions(SLFourIconSetValues IconSetType)
        {
            this.IconSetType = IconSetType;
            this.ReverseIconOrder = false;
            this.ShowIconOnly = false;

            this.IsCustomIcon = false;

            this.GreaterThanOrEqual2 = true;
            this.GreaterThanOrEqual3 = true;
            this.GreaterThanOrEqual4 = true;

            switch (IconSetType)
            {
                case SLFourIconSetValues.FourArrows:
                    this.vIcon1 = SLIconValues.RedDownArrow;
                    this.vIcon2 = SLIconValues.YellowDownInclineArrow;
                    this.vIcon3 = SLIconValues.YellowUpInclineArrow;
                    this.vIcon4 = SLIconValues.GreenUpArrow;
                    break;
                case SLFourIconSetValues.FourArrowsGray:
                    this.vIcon1 = SLIconValues.GrayDownArrow;
                    this.vIcon2 = SLIconValues.GrayDownInclineArrow;
                    this.vIcon3 = SLIconValues.GrayUpInclineArrow;
                    this.vIcon4 = SLIconValues.GrayUpArrow;
                    break;
                case SLFourIconSetValues.FourRating:
                    this.vIcon1 = SLIconValues.SignalMeterWithOneFilledBar;
                    this.vIcon2 = SLIconValues.SignalMeterWithTwoFilledBars;
                    this.vIcon3 = SLIconValues.SignalMeterWithThreeFilledBars;
                    this.vIcon4 = SLIconValues.SignalMeterWithFourFilledBars;
                    break;
                case SLFourIconSetValues.FourRedToBlack:
                    this.vIcon1 = SLIconValues.BlackCircle;
                    this.vIcon2 = SLIconValues.GrayCircle;
                    this.vIcon3 = SLIconValues.PinkCircle;
                    this.vIcon4 = SLIconValues.RedCircle;
                    break;
                case SLFourIconSetValues.FourTrafficLights:
                    this.vIcon1 = SLIconValues.BlackCircleWithBorder;
                    this.vIcon2 = SLIconValues.RedCircleWithBorder;
                    this.vIcon3 = SLIconValues.YellowCircle;
                    this.vIcon4 = SLIconValues.GreenCircle;
                    break;
            }

            this.Value2 = "25";
            this.Value3 = "50";
            this.Value4 = "75";

            this.Type2 = SLConditionalFormatRangeValues.Percent;
            this.Type3 = SLConditionalFormatRangeValues.Percent;
            this.Type4 = SLConditionalFormatRangeValues.Percent;
        }
    }
}
