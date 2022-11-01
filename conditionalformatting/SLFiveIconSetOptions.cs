using System;
using System.Collections.Generic;

namespace SpreadsheetLight
{
    /// <summary>
    /// Conditional formatting options for five icon sets.
    /// </summary>
    public class SLFiveIconSetOptions
    {
        internal SLFiveIconSetValues IconSetType { get; set; }

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

        internal SLIconValues vIcon5;
        /// <summary>
        /// The 5th icon.
        /// </summary>
        public SLIconValues Icon5
        {
            get { return vIcon5; }
            set
            {
                if (vIcon5 != value)
                {
                    vIcon5 = value;
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
        /// Specifies if values are to be greater than or equal to the 5th range value. Set to false if values are to be strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual5 { get; set; }

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
        /// The 5th range value.
        /// </summary>
        public string Value5 { get; set; }

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
        /// The conditional format type for the 5th range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type5 { get; set; }

        /// <summary>
        /// Initializes an instance of SLFiveIconSetOptions.
        /// </summary>
        /// <param name="IconSetType">The type of icon set.</param>
        public SLFiveIconSetOptions(SLFiveIconSetValues IconSetType)
        {
            this.IconSetType = IconSetType;
            this.ReverseIconOrder = false;
            this.ShowIconOnly = false;

            this.IsCustomIcon = false;

            this.GreaterThanOrEqual2 = true;
            this.GreaterThanOrEqual3 = true;
            this.GreaterThanOrEqual4 = true;
            this.GreaterThanOrEqual5 = true;

            switch (IconSetType)
            {
                case SLFiveIconSetValues.FiveArrows:
                    this.vIcon1 = SLIconValues.RedDownArrow;
                    this.vIcon2 = SLIconValues.YellowDownInclineArrow;
                    this.vIcon3 = SLIconValues.YellowSideArrow;
                    this.vIcon4 = SLIconValues.YellowUpInclineArrow;
                    this.vIcon5 = SLIconValues.GreenUpArrow;
                    break;
                case SLFiveIconSetValues.FiveArrowsGray:
                    this.vIcon1 = SLIconValues.GrayDownArrow;
                    this.vIcon2 = SLIconValues.GrayDownInclineArrow;
                    this.vIcon3 = SLIconValues.GraySideArrow;
                    this.vIcon4 = SLIconValues.GrayUpInclineArrow;
                    this.vIcon5 = SLIconValues.GrayUpArrow;
                    break;
                case SLFiveIconSetValues.FiveBoxes:
                    this.vIcon1 = SLIconValues.ZeroFilledBoxes;
                    this.vIcon2 = SLIconValues.OneFilledBox;
                    this.vIcon3 = SLIconValues.TwoFilledBoxes;
                    this.vIcon4 = SLIconValues.ThreeFilledBoxes;
                    this.vIcon5 = SLIconValues.FourFilledBoxes;
                    break;
                case SLFiveIconSetValues.FiveQuarters:
                    this.vIcon1 = SLIconValues.WhiteCircleAllWhiteQuarters;
                    this.vIcon2 = SLIconValues.CircleWithThreeWhiteQuarters;
                    this.vIcon3 = SLIconValues.CircleWithTwoWhiteQuarters;
                    this.vIcon4 = SLIconValues.CircleWithOneWhiteQuarter;
                    this.vIcon5 = SLIconValues.BlackCircle;
                    break;
                case SLFiveIconSetValues.FiveRating:
                    this.vIcon1 = SLIconValues.SignalMeterWithNoFilledBars;
                    this.vIcon2 = SLIconValues.SignalMeterWithOneFilledBar;
                    this.vIcon3 = SLIconValues.SignalMeterWithTwoFilledBars;
                    this.vIcon4 = SLIconValues.SignalMeterWithThreeFilledBars;
                    this.vIcon5 = SLIconValues.SignalMeterWithFourFilledBars;
                    break;
            }

            this.Value2 = "20";
            this.Value3 = "40";
            this.Value4 = "60";
            this.Value5 = "80";

            this.Type2 = SLConditionalFormatRangeValues.Percent;
            this.Type3 = SLConditionalFormatRangeValues.Percent;
            this.Type4 = SLConditionalFormatRangeValues.Percent;
            this.Type5 = SLConditionalFormatRangeValues.Percent;
        }
    }
}
