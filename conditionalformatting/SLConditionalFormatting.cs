using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    internal enum SLConditionalFormatVersionValue
    {
        Office = 0,
        Office2010
    }

    internal struct SLConditionalFormatOrder
    {
        internal SLConditionalFormatVersionValue Version;
        internal int? Priority;
        internal int Order;

        internal SLConditionalFormatOrder(SLConditionalFormatVersionValue Version, int? Priority, int Order)
        {
            this.Version = Version;
            this.Priority = Priority;
            this.Order = Order; // 0 by default?
        }
    }

    /// <summary>
    /// Built-in data bar types.
    /// </summary>
    public enum SLConditionalFormatDataBarValues
    {
        /// <summary>
        /// Blue data bar
        /// </summary>
        Blue = 0,
        /// <summary>
        /// Green data bar
        /// </summary>
        Green,
        /// <summary>
        /// Red data bar
        /// </summary>
        Red,
        /// <summary>
        /// Orange data bar
        /// </summary>
        Orange,
        /// <summary>
        /// Light blue data bar
        /// </summary>
        LightBlue,
        /// <summary>
        /// Purple data bar
        /// </summary>
        Purple
    }

    /// <summary>
    /// Built-in color scale types.
    /// </summary>
    public enum SLConditionalFormatColorScaleValues
    {
        /// <summary>
        /// Green - Yellow - Red color scale
        /// </summary>
        GreenYellowRed = 0,
        /// <summary>
        /// Red - Yellow - Green color scale
        /// </summary>
        RedYellowGreen,
        /// <summary>
        /// Blue - Yellow - Red color scale
        /// </summary>
        BlueYellowRed,
        /// <summary>
        /// Red - Yellow - Blue color scale
        /// </summary>
        RedYellowBlue,
        /// <summary>
        /// Green - White - Red color scale
        /// </summary>
        GreenWhiteRed, // Excel 2010
        /// <summary>
        /// Red - White - Green color scale
        /// </summary>
        RedWhiteGreen, // Excel 2010
        /// <summary>
        /// Blue - White - Red color scale
        /// </summary>
        BlueWhiteRed, // Excel 2010
        /// <summary>
        /// Red - White - Blue color scale
        /// </summary>
        RedWhiteBlue, // Excel 2010
        /// <summary>
        /// White - Red color scale
        /// </summary>
        WhiteRed, // Excel 2010
        /// <summary>
        /// Red - White color scale
        /// </summary>
        RedWhite, // Excel 2010
        /// <summary>
        /// Green - White color scale
        /// </summary>
        GreenWhite, // Excel 2010
        /// <summary>
        /// White - Green color scale
        /// </summary>
        WhiteGreen, // Excel 2010
        /// <summary>
        /// Yellow - Red color scale
        /// </summary>
        YellowRed,
        /// <summary>
        /// Red - Yellow color scale
        /// </summary>
        RedYellow,
        /// <summary>
        /// Green - Yellow color scale
        /// </summary>
        GreenYellow,
        /// <summary>
        /// Yellow - Green color scale
        /// </summary>
        YellowGreen
    }

    /// <summary>
    /// Conditional format type including minimum and maximum types.
    /// </summary>
    public enum SLConditionalFormatMinMaxValues
    {
        /// <summary>
        /// The underlying engine will assign a minimum or maximum depending on the parameter this value is used on.
        /// </summary>
        Value = 0,
        /// <summary>
        /// Number
        /// </summary>
        Number,
        /// <summary>
        /// Percent
        /// </summary>
        Percent,
        /// <summary>
        /// Formula
        /// </summary>
        Formula,
        /// <summary>
        /// Percentile
        /// </summary>
        Percentile
    }

    /// <summary>
    /// Conditional format type including minimum and maximum and automatic types.
    /// </summary>
    public enum SLConditionalFormatAutoMinMaxValues
    {
        /// <summary>
        /// The underlying engine will assign a minimum or maximum depending on the parameter this value is used on.
        /// </summary>
        Value = 0,
        /// <summary>
        /// Number
        /// </summary>
        Number,
        /// <summary>
        /// Percent
        /// </summary>
        Percent,
        /// <summary>
        /// Formula
        /// </summary>
        Formula,
        /// <summary>
        /// Percentile
        /// </summary>
        Percentile,
        /// <summary>
        /// Automatic. The underlying engine will assign a minimum or maximum depending on the parameter this value is used on. This is an Excel 2010 specific feature.
        /// </summary>
        Automatic
    }

    /// <summary>
    /// Conditional format type excluding minimum and maximum and automatic types.
    /// </summary>
    public enum SLConditionalFormatRangeValues
    {
        /// <summary>
        /// Number
        /// </summary>
        Number = 0,
        /// <summary>
        /// Percent
        /// </summary>
        Percent,
        /// <summary>
        /// Formula
        /// </summary>
        Formula,
        /// <summary>
        /// Percentile
        /// </summary>
        Percentile
    }

    internal enum SLIconSetValues
    {
        /// <summary>
        /// 5 arrows
        /// </summary>
        FiveArrows = 0,
        /// <summary>
        /// 5 arrows (gray)
        /// </summary>
        FiveArrowsGray,
        /// <summary>
        /// 5 boxes
        /// </summary>
        FiveBoxes, // Excel 2010
        /// <summary>
        /// 5 quarters
        /// </summary>
        FiveQuarters,
        /// <summary>
        /// 5 ratings
        /// </summary>
        FiveRating,
        /// <summary>
        /// 4 arrows
        /// </summary>
        FourArrows,
        /// <summary>
        /// 4 arrows (gray)
        /// </summary>
        FourArrowsGray,
        /// <summary>
        /// 4 ratings
        /// </summary>
        FourRating,
        /// <summary>
        /// 4 red To black
        /// </summary>
        FourRedToBlack,
        /// <summary>
        /// 4 traffic lights
        /// </summary>
        FourTrafficLights,
        /// <summary>
        /// 3 arrows
        /// </summary>
        ThreeArrows,
        /// <summary>
        /// 3 arrows (gray)
        /// </summary>
        ThreeArrowsGray,
        /// <summary>
        /// 3 flags
        /// </summary>
        ThreeFlags,
        /// <summary>
        /// 3 signs
        /// </summary>
        ThreeSigns,
        /// <summary>
        /// 3 stars
        /// </summary>
        ThreeStars, // Excel 2010
        /// <summary>
        /// 3 symbols circled
        /// </summary>
        ThreeSymbols,
        /// <summary>
        /// 3 symbols
        /// </summary>
        ThreeSymbols2,
        /// <summary>
        /// 3 traffic lights
        /// </summary>
        ThreeTrafficLights1,
        /// <summary>
        /// 3 traffic lights black
        /// </summary>
        ThreeTrafficLights2,
        /// <summary>
        /// 3 triangles
        /// </summary>
        ThreeTriangles // Excel 2010
    }

    /// <summary>
    /// Icon set type for five icons.
    /// </summary>
    public enum SLFiveIconSetValues
    {
        /// <summary>
        /// 5 arrows
        /// </summary>
        FiveArrows = 0,
        /// <summary>
        /// 5 arrows (gray)
        /// </summary>
        FiveArrowsGray,
        /// <summary>
        /// 5 boxes. This is an Excel 2010 specific feature.
        /// </summary>
        FiveBoxes, // Excel 2010
        /// <summary>
        /// 5 quarters
        /// </summary>
        FiveQuarters,
        /// <summary>
        /// 5 ratings
        /// </summary>
        FiveRating
    }

    /// <summary>
    /// Icon set type for four icons.
    /// </summary>
    public enum SLFourIconSetValues
    {
        /// <summary>
        /// 4 arrows
        /// </summary>
        FourArrows = 0,
        /// <summary>
        /// 4 arrows (gray)
        /// </summary>
        FourArrowsGray,
        /// <summary>
        /// 4 ratings
        /// </summary>
        FourRating,
        /// <summary>
        /// 4 red To black
        /// </summary>
        FourRedToBlack,
        /// <summary>
        /// 4 traffic lights
        /// </summary>
        FourTrafficLights
    }

    /// <summary>
    /// Icon set type for three icons.
    /// </summary>
    public enum SLThreeIconSetValues
    {
        /// <summary>
        /// 3 arrows
        /// </summary>
        ThreeArrows = 0,
        /// <summary>
        /// 3 arrows (gray)
        /// </summary>
        ThreeArrowsGray,
        /// <summary>
        /// 3 flags
        /// </summary>
        ThreeFlags,
        /// <summary>
        /// 3 signs
        /// </summary>
        ThreeSigns,
        /// <summary>
        /// 3 stars. This is an Excel 2010 specific feature.
        /// </summary>
        ThreeStars, // Excel 2010
        /// <summary>
        /// 3 symbols circled
        /// </summary>
        ThreeSymbols,
        /// <summary>
        /// 3 symbols
        /// </summary>
        ThreeSymbols2,
        /// <summary>
        /// 3 traffic lights
        /// </summary>
        ThreeTrafficLights1,
        /// <summary>
        /// 3 traffic lights black
        /// </summary>
        ThreeTrafficLights2,
        /// <summary>
        /// 3 triangles. This is an Excel 2010 specific feature.
        /// </summary>
        ThreeTriangles // Excel 2010
    }

    /// <summary>
    /// Icon types.
    /// </summary>
    public enum SLIconValues
    {
        /// <summary>
        /// No icon.
        /// </summary>
        NoIcon = 0,
        /// <summary>
        /// Green up arrow.
        /// </summary>
        GreenUpArrow,
        /// <summary>
        /// Yellow side arrow.
        /// </summary>
        YellowSideArrow,
        /// <summary>
        /// Red down arrow.
        /// </summary>
        RedDownArrow,
        /// <summary>
        /// Gray up arrow.
        /// </summary>
        GrayUpArrow,
        /// <summary>
        /// Gray side arrow.
        /// </summary>
        GraySideArrow,
        /// <summary>
        /// Gray down arrow.
        /// </summary>
        GrayDownArrow,
        /// <summary>
        /// Green flag.
        /// </summary>
        GreenFlag,
        /// <summary>
        /// Yellow flag.
        /// </summary>
        YellowFlag,
        /// <summary>
        /// Red flag.
        /// </summary>
        RedFlag,
        /// <summary>
        /// Green circle.
        /// </summary>
        GreenCircle,
        /// <summary>
        /// Yellow circle.
        /// </summary>
        YellowCircle,
        /// <summary>
        /// Red circle with border.
        /// </summary>
        RedCircleWithBorder,
        /// <summary>
        /// Black circle with border.
        /// </summary>
        BlackCircleWithBorder,
        /// <summary>
        /// Green traffic light.
        /// </summary>
        GreenTrafficLight,
        /// <summary>
        /// Yellow traffic light.
        /// </summary>
        YellowTrafficLight,
        /// <summary>
        /// Red traffic light.
        /// </summary>
        RedTrafficLight,
        /// <summary>
        /// Yellow triangle.
        /// </summary>
        YellowTriangle,
        /// <summary>
        /// Red diamond.
        /// </summary>
        RedDiamond,
        /// <summary>
        /// Green check symbol.
        /// </summary>
        GreenCheckSymbol,
        /// <summary>
        /// Yellow exclamation symbol.
        /// </summary>
        YellowExclamationSymbol,
        /// <summary>
        /// Red cross symbol.
        /// </summary>
        RedCrossSymbol,
        /// <summary>
        /// Green check.
        /// </summary>
        GreenCheck,
        /// <summary>
        /// Yellow exclamation.
        /// </summary>
        YellowExclamation,
        /// <summary>
        /// Red cross.
        /// </summary>
        RedCross,
        /// <summary>
        /// Yellow up incline arrow.
        /// </summary>
        YellowUpInclineArrow,
        /// <summary>
        /// Yellow down incline arrow.
        /// </summary>
        YellowDownInclineArrow,
        /// <summary>
        /// Gray up incline arrow.
        /// </summary>
        GrayUpInclineArrow,
        /// <summary>
        /// Gray down incline arrow.
        /// </summary>
        GrayDownInclineArrow,
        /// <summary>
        /// Red circle.
        /// </summary>
        RedCircle,
        /// <summary>
        /// Pink circle.
        /// </summary>
        PinkCircle,
        /// <summary>
        /// Gray circle.
        /// </summary>
        GrayCircle,
        /// <summary>
        /// Black circle.
        /// </summary>
        BlackCircle,
        /// <summary>
        /// Circle with one white quarter.
        /// </summary>
        CircleWithOneWhiteQuarter,
        /// <summary>
        /// Circle with two white quarters.
        /// </summary>
        CircleWithTwoWhiteQuarters,
        /// <summary>
        /// Circle with three white quarters.
        /// </summary>
        CircleWithThreeWhiteQuarters,
        /// <summary>
        /// White circle (all white quarters).
        /// </summary>
        WhiteCircleAllWhiteQuarters,
        /// <summary>
        /// Signal meter with no filled bars.
        /// </summary>
        SignalMeterWithNoFilledBars,
        /// <summary>
        /// Signal meter with one filled bar.
        /// </summary>
        SignalMeterWithOneFilledBar,
        /// <summary>
        /// Signal meter with two filled bars.
        /// </summary>
        SignalMeterWithTwoFilledBars,
        /// <summary>
        /// Signal meter with three filled bars.
        /// </summary>
        SignalMeterWithThreeFilledBars,
        /// <summary>
        /// Signal meter with four filled bars.
        /// </summary>
        SignalMeterWithFourFilledBars,
        /// <summary>
        /// Gold star.
        /// </summary>
        GoldStar,
        /// <summary>
        /// Half gold star.
        /// </summary>
        HalfGoldStar,
        /// <summary>
        /// Silver star.
        /// </summary>
        SilverStar,
        /// <summary>
        /// Green up triangle.
        /// </summary>
        GreenUpTriangle,
        /// <summary>
        /// Yellow dash.
        /// </summary>
        YellowDash,
        /// <summary>
        /// Red down triangle.
        /// </summary>
        RedDownTriangle,
        /// <summary>
        /// Four filled boxes.
        /// </summary>
        FourFilledBoxes,
        /// <summary>
        /// Three filled boxes.
        /// </summary>
        ThreeFilledBoxes,
        /// <summary>
        /// Two filled boxes.
        /// </summary>
        TwoFilledBoxes,
        /// <summary>
        /// One filled box.
        /// </summary>
        OneFilledBox,
        /// <summary>
        /// Zero filled boxes.
        /// </summary>
        ZeroFilledBoxes
    }

    /// <summary>
    /// Built-in cell highlighting styles
    /// </summary>
    public enum SLHighlightCellsStyleValues
    {
        /// <summary>
        /// Light red background fill with dark red text
        /// </summary>
        LightRedFillWithDarkRedText = 0,
        /// <summary>
        /// Yellow background fill with dark yellow text
        /// </summary>
        YellowFillWithDarkYellowText,
        /// <summary>
        /// Green background fill with dark green text
        /// </summary>
        GreenFillWithDarkGreenText,
        /// <summary>
        /// Light red background fill
        /// </summary>
        LightRedFill,
        /// <summary>
        /// Red text
        /// </summary>
        RedText,
        /// <summary>
        /// Red borders
        /// </summary>
        RedBorder
    }

    /// <summary>
    /// Options on the average value for the selected range.
    /// </summary>
    public enum SLHighlightCellsAboveAverageValues
    {
        /// <summary>
        /// Above the average
        /// </summary>
        Above = 0,
        /// <summary>
        /// Below the average
        /// </summary>
        Below,
        /// <summary>
        /// Equal to or above the average
        /// </summary>
        EqualOrAbove,
        /// <summary>
        /// Equal to or below the average
        /// </summary>
        EqualOrBelow,
        /// <summary>
        /// 1 standard deviation above the average
        /// </summary>
        OneStdDevAbove,
        /// <summary>
        /// 1 standard deviation below the average
        /// </summary>
        OneStdDevBelow,
        /// <summary>
        /// 2 standard deviations above the average
        /// </summary>
        TwoStdDevAbove,
        /// <summary>
        /// 2 standard deviations below the average
        /// </summary>
        TwoStdDevBelow,
        /// <summary>
        /// 3 standard deviations above the average
        /// </summary>
        ThreeStdDevAbove,
        /// <summary>
        /// 3 standard deviations below the average
        /// </summary>
        ThreeStdDevBelow
    }

    /// <summary>
    /// Encapsulates properties and methods for conditional formatting. This simulates the DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting class.
    /// </summary>
    public class SLConditionalFormatting
    {
        // Conditional formatting doesn't need the theme or indexed colours.

        internal List<SLConditionalFormattingRule> Rules { get; set; }
        internal bool Pivot { get; set; }
        internal List<SLCellPointRange> SequenceOfReferences { get; set; }

        /// <summary>
        /// Initializes an instance of SLConditionalFormatting, given cell references of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartCellReference">The cell reference of the start cell of the cell range to be conditionally formatted, such as "A1". This is typically the top-left cell.</param>
        /// <param name="EndCellReference">The cell reference of the end cell of the cell range to be conditionally formatted, such as "A1". This is typically the bottom-right cell.</param>
        public SLConditionalFormatting(string StartCellReference, string EndCellReference)
        {
            int iStartRowIndex = -1;
            int iStartColumnIndex = -1;
            int iEndRowIndex = -1;
            int iEndColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex, out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            this.InitialiseNewConditionalFormatting(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        /// Initializes an instance of SLConditionalFormatting, given row and column indices of opposite cells in a cell range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        public SLConditionalFormatting(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            this.InitialiseNewConditionalFormatting(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex);
        }

        internal SLConditionalFormatting()
        {
            this.SetAllNull();
        }

        private void InitialiseNewConditionalFormatting(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            this.SetAllNull();

            this.SequenceOfReferences.Add(new SLCellPointRange(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex));
        }

        private void SetAllNull()
        {
            this.Rules = new List<SLConditionalFormattingRule>();
            this.Pivot = false;
            this.SequenceOfReferences = new List<SLCellPointRange>();
        }

        private void AppendRule(SLConditionalFormattingRule cfr)
        {
            if (this.Rules.Count > 0)
            {
                int index = this.Rules.Count - 1;
                // This follows the Excel behaviour.
                // If the last rule is of the same type, then the last rule
                // is overwritten with the newly given rule.
                if (this.Rules[index].Type == cfr.Type)
                {
                    this.Rules[index] = cfr;
                }
                else
                {
                    this.Rules.Add(cfr);
                }
            }
            else
            {
                this.Rules.Add(cfr);
            }
        }

        /// <summary>
        /// Set a data bar formatting with built-in types.
        /// </summary>
        /// <param name="DataBar">A built-in data bar type.</param>
        public void SetDataBar(SLConditionalFormatDataBarValues DataBar)
        {
            SLDataBarOptions dbo = new SLDataBarOptions(DataBar, false);
            this.SetCustomDataBar(dbo);
        }

        /// <summary>
        /// Set a custom data bar formatting.
        /// </summary>
        /// <param name="ShowBarOnly">True to show only the data bar. False to show both data bar and value.</param>
        /// <param name="MinLength">The minimum length of the data bar as a percentage of the cell width. The default value is 10.</param>
        /// <param name="MaxLength">The maximum length of the data bar as a percentage of the cell width. The default value is 90.</param>
        /// <param name="ShortestBarType">The conditional format type for the shortest bar.</param>
        /// <param name="ShortestBarValue">The value for the shortest bar. If <paramref name="ShortestBarType"/> is Value, you can just set this to "0".</param>
        /// <param name="LongestBarType">The conditional format type for the longest bar.</param>
        /// <param name="LongestBarValue">The value for the longest bar. If <paramref name="LongestBarType"/> is Value, you can just set this to "0".</param>
        /// <param name="BarColor">The color of the data bar.</param>
        public void SetCustomDataBar(bool ShowBarOnly, uint MinLength, uint MaxLength, SLConditionalFormatMinMaxValues ShortestBarType, string ShortestBarValue, SLConditionalFormatMinMaxValues LongestBarType, string LongestBarValue, System.Drawing.Color BarColor)
        {
            SLDataBarOptions dbo = new SLDataBarOptions(false);
            dbo.ShowBarOnly = ShowBarOnly;
            dbo.MinLength = MinLength;
            dbo.MaxLength = MaxLength;
            dbo.MinimumType = this.TranslateMinMaxValues(ShortestBarType);
            dbo.MinimumValue = ShortestBarValue;
            dbo.MaximumType = this.TranslateMinMaxValues(LongestBarType);
            dbo.MaximumValue = LongestBarValue;
            dbo.FillColor.Color = BarColor;

            this.SetCustomDataBar(dbo);
        }

        /// <summary>
        /// Set a custom data bar formatting.
        /// </summary>
        /// <param name="ShowBarOnly">True to show only the data bar. False to show both data bar and value.</param>
        /// <param name="MinLength">The minimum length of the data bar as a percentage of the cell width. The default value is 10.</param>
        /// <param name="MaxLength">The maximum length of the data bar as a percentage of the cell width. The default value is 90.</param>
        /// <param name="ShortestBarType">The conditional format type for the shortest bar.</param>
        /// <param name="ShortestBarValue">The value for the shortest bar. If <paramref name="ShortestBarType"/> is Value, you can just set this to "0".</param>
        /// <param name="LongestBarType">The conditional format type for the longest bar.</param>
        /// <param name="LongestBarValue">The value for the longest bar. If <paramref name="LongestBarType"/> is Value, you can just set this to "0".</param>
        /// <param name="BarColor">The theme color to be used for the data bar.</param>
        public void SetCustomDataBar(bool ShowBarOnly, uint MinLength, uint MaxLength, SLConditionalFormatMinMaxValues ShortestBarType, string ShortestBarValue, SLConditionalFormatMinMaxValues LongestBarType, string LongestBarValue, SLThemeColorIndexValues BarColor)
        {
            SLDataBarOptions dbo = new SLDataBarOptions(false);
            dbo.ShowBarOnly = ShowBarOnly;
            dbo.MinLength = MinLength;
            dbo.MaxLength = MaxLength;
            dbo.MinimumType = this.TranslateMinMaxValues(ShortestBarType);
            dbo.MinimumValue = ShortestBarValue;
            dbo.MaximumType = this.TranslateMinMaxValues(LongestBarType);
            dbo.MaximumValue = LongestBarValue;
            dbo.FillColor.SetThemeColor(BarColor);

            this.SetCustomDataBar(dbo);
        }

        /// <summary>
        /// Set a custom data bar formatting.
        /// </summary>
        /// <param name="ShowBarOnly">True to show only the data bar. False to show both data bar and value.</param>
        /// <param name="MinLength">The minimum length of the data bar as a percentage of the cell width. The default value is 10.</param>
        /// <param name="MaxLength">The maximum length of the data bar as a percentage of the cell width. The default value is 90.</param>
        /// <param name="ShortestBarType">The conditional format type for the shortest bar.</param>
        /// <param name="ShortestBarValue">The value for the shortest bar. If <paramref name="ShortestBarType"/> is Value, you can just set this to "0".</param>
        /// <param name="LongestBarType">The conditional format type for the longest bar.</param>
        /// <param name="LongestBarValue">The value for the longest bar. If <paramref name="LongestBarType"/> is Value, you can just set this to "0".</param>
        /// <param name="BarColor">The theme color to be used for the data bar.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetCustomDataBar(bool ShowBarOnly, uint MinLength, uint MaxLength, SLConditionalFormatMinMaxValues ShortestBarType, string ShortestBarValue, SLConditionalFormatMinMaxValues LongestBarType, string LongestBarValue, SLThemeColorIndexValues BarColor, double Tint)
        {
            SLDataBarOptions dbo = new SLDataBarOptions(false);
            dbo.ShowBarOnly = ShowBarOnly;
            dbo.MinLength = MinLength;
            dbo.MaxLength = MaxLength;
            dbo.MinimumType = this.TranslateMinMaxValues(ShortestBarType);
            dbo.MinimumValue = ShortestBarValue;
            dbo.MaximumType = this.TranslateMinMaxValues(LongestBarType);
            dbo.MaximumValue = LongestBarValue;
            dbo.FillColor.SetThemeColor(BarColor, Tint);

            this.SetCustomDataBar(dbo);
        }

        /// <summary>
        /// Set a custom data bar formatting.
        /// </summary>
        /// <param name="Options">Data bar options.</param>
        public void SetCustomDataBar(SLDataBarOptions Options)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.DataBar;

            cfr.DataBar.Is2010 = Options.Is2010;

            cfr.DataBar.MinimumType = Options.MinimumType;
            cfr.DataBar.MinimumValue = Options.MinimumValue;
            cfr.DataBar.MaximumType = Options.MaximumType;
            cfr.DataBar.MaximumValue = Options.MaximumValue;

            cfr.DataBar.Color = Options.FillColor.Clone();
            cfr.DataBar.BorderColor = Options.BorderColor.Clone();
            cfr.DataBar.NegativeFillColor = Options.NegativeFillColor.Clone();
            cfr.DataBar.NegativeBorderColor = Options.NegativeBorderColor.Clone();
            cfr.DataBar.AxisColor = Options.AxisColor.Clone();

            cfr.DataBar.MinLength = Options.MinLength;
            cfr.DataBar.MaxLength = Options.MaxLength;
            cfr.DataBar.ShowValue = !Options.ShowBarOnly;
            cfr.DataBar.Border = Options.Border;
            cfr.DataBar.Gradient = Options.Gradient;
            cfr.DataBar.Direction = Options.Direction;
            cfr.DataBar.NegativeBarColorSameAsPositive = Options.NegativeBarColorSameAsPositive;
            cfr.DataBar.NegativeBarBorderColorSameAsPositive = Options.NegativeBarBorderColorSameAsPositive;
            cfr.DataBar.AxisPosition = Options.AxisPosition;

            cfr.HasDataBar = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Set a color scale formatting with built-in types.
        /// </summary>
        /// <param name="ColorScale">A built-in color scale type.</param>
        public void SetColorScale(SLConditionalFormatColorScaleValues ColorScale)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.ColorScale;
            cfr.ColorScale.Cfvos.Add(new SLConditionalFormatValueObject()
            {
                Type = ConditionalFormatValueObjectValues.Min //Excel 2010 omits this: , Val = "0"
            });
            if (ColorScale == SLConditionalFormatColorScaleValues.GreenYellowRed
                || ColorScale == SLConditionalFormatColorScaleValues.RedYellowGreen
                || ColorScale == SLConditionalFormatColorScaleValues.BlueYellowRed
                || ColorScale == SLConditionalFormatColorScaleValues.RedYellowBlue
                || ColorScale == SLConditionalFormatColorScaleValues.GreenWhiteRed
                || ColorScale == SLConditionalFormatColorScaleValues.RedWhiteGreen
                || ColorScale == SLConditionalFormatColorScaleValues.BlueWhiteRed
                || ColorScale == SLConditionalFormatColorScaleValues.RedWhiteBlue)
            {
                cfr.ColorScale.Cfvos.Add(new SLConditionalFormatValueObject()
                {
                    Type = ConditionalFormatValueObjectValues.Percentile,
                    Val = "50"
                });
            }
            cfr.ColorScale.Cfvos.Add(new SLConditionalFormatValueObject()
            {
                Type = ConditionalFormatValueObjectValues.Max //Excel 2010 omits this: , Val = "0"
            });

            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            switch (ColorScale)
            {
                case SLConditionalFormatColorScaleValues.GreenYellowRed:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEB, 0x84) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xBE, 0x7B) });
                    break;
                case SLConditionalFormatColorScaleValues.RedYellowGreen:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xBE, 0x7B) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEB, 0x84) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    break;
                case SLConditionalFormatColorScaleValues.BlueYellowRed:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEB, 0x84) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x5A, 0x8A, 0xC6) });
                    break;
                case SLConditionalFormatColorScaleValues.RedYellowBlue:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x5A, 0x8A, 0xC6) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEB, 0x84) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    break;
                case SLConditionalFormatColorScaleValues.GreenWhiteRed:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFC, 0xFC, 0xFF) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xBE, 0x7B) });
                    break;
                case SLConditionalFormatColorScaleValues.RedWhiteGreen:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xBE, 0x7B) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFC, 0xFC, 0xFF) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    break;
                case SLConditionalFormatColorScaleValues.BlueWhiteRed:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFC, 0xFC, 0xFF) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x5A, 0x8A, 0xC6) });
                    break;
                case SLConditionalFormatColorScaleValues.RedWhiteBlue:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x5A, 0x8A, 0xC6) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFC, 0xFC, 0xFF) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    break;
                case SLConditionalFormatColorScaleValues.WhiteRed:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFC, 0xFC, 0xFF) });
                    break;
                case SLConditionalFormatColorScaleValues.RedWhite:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFC, 0xFC, 0xFF) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xF8, 0x69, 0x6B) });
                    break;
                case SLConditionalFormatColorScaleValues.GreenWhite:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFC, 0xFC, 0xFF) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xBE, 0x7B) });
                    break;
                case SLConditionalFormatColorScaleValues.WhiteGreen:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xBE, 0x7B) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFC, 0xFC, 0xFF) });
                    break;
                case SLConditionalFormatColorScaleValues.YellowRed:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x71, 0x28) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEF, 0x9C) });
                    break;
                case SLConditionalFormatColorScaleValues.RedYellow:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEF, 0x9C) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0x71, 0x28) });
                    break;
                case SLConditionalFormatColorScaleValues.GreenYellow:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEF, 0x9C) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xBE, 0x7B) });
                    break;
                case SLConditionalFormatColorScaleValues.YellowGreen:
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0x63, 0xBE, 0x7B) });
                    cfr.ColorScale.Colors.Add(new SLColor(listempty, listempty) { Color = System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEF, 0x9C) });
                    break;
            }
            
            cfr.HasColorScale = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Set a custom 2-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The color for the minimum.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The color for the maximum.</param>
        public void SetCustom2ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, System.Drawing.Color MinColor,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, System.Drawing.Color MaxColor)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            minclr.Color = MinColor;
            SLColor maxclr = new SLColor(listempty, listempty);
            maxclr.Color = MaxColor;

            SLColor midclr = new SLColor(listempty, listempty);
            this.SetCustomColorScale(MinType, MinValue, minclr, false, SLConditionalFormatRangeValues.Percentile, "", midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 2-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The color for the minimum.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The theme color for the maximum.</param>
        /// <param name="MaxColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetCustom2ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, System.Drawing.Color MinColor,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, SLThemeColorIndexValues MaxColor, double MaxColorTint)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            minclr.Color = MinColor;
            SLColor maxclr = new SLColor(listempty, listempty);
            if (MaxColorTint == 0) maxclr.SetThemeColor(MaxColor);
            else maxclr.SetThemeColor(MaxColor, MaxColorTint);

            SLColor midclr = new SLColor(listempty, listempty);
            this.SetCustomColorScale(MinType, MinValue, minclr, false, SLConditionalFormatRangeValues.Percentile, "", midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 2-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The theme color for the minimum.</param>
        /// <param name="MinColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The color for the maximum.</param>
        public void SetCustom2ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, SLThemeColorIndexValues MinColor, double MinColorTint,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, System.Drawing.Color MaxColor)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            if (MinColorTint == 0) minclr.SetThemeColor(MinColor);
            else minclr.SetThemeColor(MinColor, MinColorTint);
            SLColor maxclr = new SLColor(listempty, listempty);
            maxclr.Color = MaxColor;

            SLColor midclr = new SLColor(listempty, listempty);
            this.SetCustomColorScale(MinType, MinValue, minclr, false, SLConditionalFormatRangeValues.Percentile, "", midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 2-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The theme color for the minimum.</param>
        /// <param name="MinColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The theme color for the maximum.</param>
        /// <param name="MaxColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetCustom2ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, SLThemeColorIndexValues MinColor, double MinColorTint,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, SLThemeColorIndexValues MaxColor, double MaxColorTint)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            if (MinColorTint == 0) minclr.SetThemeColor(MinColor);
            else minclr.SetThemeColor(MinColor, MinColorTint);
            SLColor maxclr = new SLColor(listempty, listempty);
            if (MaxColorTint == 0) maxclr.SetThemeColor(MaxColor);
            else maxclr.SetThemeColor(MaxColor, MaxColorTint);

            SLColor midclr = new SLColor(listempty, listempty);
            this.SetCustomColorScale(MinType, MinValue, minclr, false, SLConditionalFormatRangeValues.Percentile, "", midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 3-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The color for the minimum.</param>
        /// <param name="MidPointType">The conditional format type for the midpoint.</param>
        /// <param name="MidPointValue">The value for the midpoint.</param>
        /// <param name="MidPointColor">The color for the midpoint.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The color for the maximum.</param>
        public void SetCustom3ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, System.Drawing.Color MinColor,
            SLConditionalFormatRangeValues MidPointType, string MidPointValue, System.Drawing.Color MidPointColor,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, System.Drawing.Color MaxColor)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            minclr.Color = MinColor;
            SLColor maxclr = new SLColor(listempty, listempty);
            maxclr.Color = MaxColor;

            SLColor midclr = new SLColor(listempty, listempty);
            midclr.Color = MidPointColor;
            this.SetCustomColorScale(MinType, MinValue, minclr, true, MidPointType, MidPointValue, midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 3-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The color for the minimum.</param>
        /// <param name="MidPointType">The conditional format type for the midpoint.</param>
        /// <param name="MidPointValue">The value for the midpoint.</param>
        /// <param name="MidPointColor">The color for the midpoint.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The theme color for the maximum.</param>
        /// <param name="MaxColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetCustom3ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, System.Drawing.Color MinColor,
            SLConditionalFormatRangeValues MidPointType, string MidPointValue, System.Drawing.Color MidPointColor,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, SLThemeColorIndexValues MaxColor, double MaxColorTint)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            minclr.Color = MinColor;
            SLColor maxclr = new SLColor(listempty, listempty);
            if (MaxColorTint == 0) maxclr.SetThemeColor(MaxColor);
            else maxclr.SetThemeColor(MaxColor, MaxColorTint);

            SLColor midclr = new SLColor(listempty, listempty);
            midclr.Color = MidPointColor;
            this.SetCustomColorScale(MinType, MinValue, minclr, true, MidPointType, MidPointValue, midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 3-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The color for the minimum.</param>
        /// <param name="MidPointType">The conditional format type for the midpoint.</param>
        /// <param name="MidPointValue">The value for the midpoint.</param>
        /// <param name="MidPointColor">The theme color for the midpoint.</param>
        /// <param name="MidPointColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The color for the maximum.</param>
        public void SetCustom3ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, System.Drawing.Color MinColor,
            SLConditionalFormatRangeValues MidPointType, string MidPointValue, SLThemeColorIndexValues MidPointColor, double MidPointColorTint,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, System.Drawing.Color MaxColor)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            minclr.Color = MinColor;
            SLColor maxclr = new SLColor(listempty, listempty);
            maxclr.Color = MaxColor;

            SLColor midclr = new SLColor(listempty, listempty);
            if (MidPointColorTint == 0) midclr.SetThemeColor(MidPointColor);
            else midclr.SetThemeColor(MidPointColor, MidPointColorTint);
            this.SetCustomColorScale(MinType, MinValue, minclr, true, MidPointType, MidPointValue, midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 3-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The color for the minimum.</param>
        /// <param name="MidPointType">The conditional format type for the midpoint.</param>
        /// <param name="MidPointValue">The value for the midpoint.</param>
        /// <param name="MidPointColor">The theme color for the midpoint.</param>
        /// <param name="MidPointColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The theme color for the maximum.</param>
        /// <param name="MaxColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetCustom3ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, System.Drawing.Color MinColor,
            SLConditionalFormatRangeValues MidPointType, string MidPointValue, SLThemeColorIndexValues MidPointColor, double MidPointColorTint,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, SLThemeColorIndexValues MaxColor, double MaxColorTint)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            minclr.Color = MinColor;
            SLColor maxclr = new SLColor(listempty, listempty);
            if (MaxColorTint == 0) maxclr.SetThemeColor(MaxColor);
            else maxclr.SetThemeColor(MaxColor, MaxColorTint);

            SLColor midclr = new SLColor(listempty, listempty);
            if (MidPointColorTint == 0) midclr.SetThemeColor(MidPointColor);
            else midclr.SetThemeColor(MidPointColor, MidPointColorTint);
            this.SetCustomColorScale(MinType, MinValue, minclr, true, MidPointType, MidPointValue, midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 3-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The theme color for the minimum.</param>
        /// <param name="MinColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MidPointType">The conditional format type for the midpoint.</param>
        /// <param name="MidPointValue">The value for the midpoint.</param>
        /// <param name="MidPointColor">The color for the midpoint.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The color for the maximum.</param>
        public void SetCustom3ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, SLThemeColorIndexValues MinColor, double MinColorTint,
            SLConditionalFormatRangeValues MidPointType, string MidPointValue, System.Drawing.Color MidPointColor,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, System.Drawing.Color MaxColor)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            if (MinColorTint == 0) minclr.SetThemeColor(MinColor);
            else minclr.SetThemeColor(MinColor, MinColorTint);
            SLColor maxclr = new SLColor(listempty, listempty);
            maxclr.Color = MaxColor;

            SLColor midclr = new SLColor(listempty, listempty);
            midclr.Color = MidPointColor;
            this.SetCustomColorScale(MinType, MinValue, minclr, true, MidPointType, MidPointValue, midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 3-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The theme color for the minimum.</param>
        /// <param name="MinColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MidPointType">The conditional format type for the midpoint.</param>
        /// <param name="MidPointValue">The value for the midpoint.</param>
        /// <param name="MidPointColor">The color for the midpoint.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The theme color for the maximum.</param>
        /// <param name="MaxColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetCustom3ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, SLThemeColorIndexValues MinColor, double MinColorTint,
            SLConditionalFormatRangeValues MidPointType, string MidPointValue, System.Drawing.Color MidPointColor,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, SLThemeColorIndexValues MaxColor, double MaxColorTint)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            if (MinColorTint == 0) minclr.SetThemeColor(MinColor);
            else minclr.SetThemeColor(MinColor, MinColorTint);
            SLColor maxclr = new SLColor(listempty, listempty);
            if (MaxColorTint == 0) maxclr.SetThemeColor(MaxColor);
            else maxclr.SetThemeColor(MaxColor, MaxColorTint);

            SLColor midclr = new SLColor(listempty, listempty);
            midclr.Color = MidPointColor;
            this.SetCustomColorScale(MinType, MinValue, minclr, true, MidPointType, MidPointValue, midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 3-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The theme color for the minimum.</param>
        /// <param name="MinColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MidPointType">The conditional format type for the midpoint.</param>
        /// <param name="MidPointValue">The value for the midpoint.</param>
        /// <param name="MidPointColor">The theme color for the midpoint.</param>
        /// <param name="MidPointColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The color for the maximum.</param>
        public void SetCustom3ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, SLThemeColorIndexValues MinColor, double MinColorTint,
            SLConditionalFormatRangeValues MidPointType, string MidPointValue, SLThemeColorIndexValues MidPointColor, double MidPointColorTint,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, System.Drawing.Color MaxColor)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            if (MinColorTint == 0) minclr.SetThemeColor(MinColor);
            else minclr.SetThemeColor(MinColor, MinColorTint);
            SLColor maxclr = new SLColor(listempty, listempty);
            maxclr.Color = MaxColor;

            SLColor midclr = new SLColor(listempty, listempty);
            if (MidPointColorTint == 0) midclr.SetThemeColor(MidPointColor);
            else midclr.SetThemeColor(MidPointColor, MidPointColorTint);
            this.SetCustomColorScale(MinType, MinValue, minclr, true, MidPointType, MidPointValue, midclr, MaxType, MaxValue, maxclr);
        }

        /// <summary>
        /// Set a custom 3-color scale.
        /// </summary>
        /// <param name="MinType">The conditional format type for the minimum.</param>
        /// <param name="MinValue">The value for the minimum. If <paramref name="MinType"/> is Value, you can just set this to "0".</param>
        /// <param name="MinColor">The theme color for the minimum.</param>
        /// <param name="MinColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MidPointType">The conditional format type for the midpoint.</param>
        /// <param name="MidPointValue">The value for the midpoint.</param>
        /// <param name="MidPointColor">The theme color for the midpoint.</param>
        /// <param name="MidPointColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="MaxType">The conditional format type for the maximum.</param>
        /// <param name="MaxValue">The value for the maximum. If <paramref name="MaxType"/> is Value, you can just set this to "0".</param>
        /// <param name="MaxColor">The theme color for the maximum.</param>
        /// <param name="MaxColorTint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        public void SetCustom3ColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, SLThemeColorIndexValues MinColor, double MinColorTint,
            SLConditionalFormatRangeValues MidPointType, string MidPointValue, SLThemeColorIndexValues MidPointColor, double MidPointColorTint,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, SLThemeColorIndexValues MaxColor, double MaxColorTint)
        {
            List<System.Drawing.Color> listempty = new List<System.Drawing.Color>();
            SLColor minclr = new SLColor(listempty, listempty);
            if (MinColorTint == 0) minclr.SetThemeColor(MinColor);
            else minclr.SetThemeColor(MinColor, MinColorTint);
            SLColor maxclr = new SLColor(listempty, listempty);
            if (MaxColorTint == 0) maxclr.SetThemeColor(MaxColor);
            else maxclr.SetThemeColor(MaxColor, MaxColorTint);

            SLColor midclr = new SLColor(listempty, listempty);
            if (MidPointColorTint == 0) midclr.SetThemeColor(MidPointColor);
            else midclr.SetThemeColor(MidPointColor, MidPointColorTint);
            this.SetCustomColorScale(MinType, MinValue, minclr, true, MidPointType, MidPointValue, midclr, MaxType, MaxValue, maxclr);
        }

        private void SetCustomColorScale(SLConditionalFormatMinMaxValues MinType, string MinValue, SLColor MinColor,
            bool HasMidPoint, SLConditionalFormatRangeValues MidPointType, string MidPointValue, SLColor MidPointColor,
            SLConditionalFormatMinMaxValues MaxType, string MaxValue, SLColor MaxColor)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.ColorScale;

            SLConditionalFormatValueObject cfvo;

            cfvo = new SLConditionalFormatValueObject();
            switch (MinType)
            {
                case SLConditionalFormatMinMaxValues.Value:
                    cfvo.Type = ConditionalFormatValueObjectValues.Min;
                    break;
                case SLConditionalFormatMinMaxValues.Number:
                    cfvo.Type = ConditionalFormatValueObjectValues.Number;
                    break;
                case SLConditionalFormatMinMaxValues.Percent:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percent;
                    break;
                case SLConditionalFormatMinMaxValues.Formula:
                    cfvo.Type = ConditionalFormatValueObjectValues.Formula;
                    break;
                case SLConditionalFormatMinMaxValues.Percentile:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percentile;
                    break;
            }
            cfvo.Val = MinValue;
            cfr.ColorScale.Cfvos.Add(cfvo);
            cfr.ColorScale.Colors.Add(MinColor.Clone());

            if (HasMidPoint)
            {
                cfvo = new SLConditionalFormatValueObject();
                switch (MidPointType)
                {
                    case SLConditionalFormatRangeValues.Number:
                        cfvo.Type = ConditionalFormatValueObjectValues.Number;
                        break;
                    case SLConditionalFormatRangeValues.Percent:
                        cfvo.Type = ConditionalFormatValueObjectValues.Percent;
                        break;
                    case SLConditionalFormatRangeValues.Formula:
                        cfvo.Type = ConditionalFormatValueObjectValues.Formula;
                        break;
                    case SLConditionalFormatRangeValues.Percentile:
                        cfvo.Type = ConditionalFormatValueObjectValues.Percentile;
                        break;
                }
                cfvo.Val = MidPointValue;
                cfr.ColorScale.Cfvos.Add(cfvo);
                cfr.ColorScale.Colors.Add(MidPointColor.Clone());
            }

            cfvo = new SLConditionalFormatValueObject();
            switch (MaxType)
            {
                case SLConditionalFormatMinMaxValues.Value:
                    cfvo.Type = ConditionalFormatValueObjectValues.Max;
                    break;
                case SLConditionalFormatMinMaxValues.Number:
                    cfvo.Type = ConditionalFormatValueObjectValues.Number;
                    break;
                case SLConditionalFormatMinMaxValues.Percent:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percent;
                    break;
                case SLConditionalFormatMinMaxValues.Formula:
                    cfvo.Type = ConditionalFormatValueObjectValues.Formula;
                    break;
                case SLConditionalFormatMinMaxValues.Percentile:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percentile;
                    break;
            }
            cfvo.Val = MaxValue;
            cfr.ColorScale.Cfvos.Add(cfvo);
            cfr.ColorScale.Colors.Add(MaxColor.Clone());

            cfr.HasColorScale = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Set an icon set formatting with built-in types.
        /// </summary>
        /// <param name="IconSetType">A built-in icon set type.</param>
        public void SetIconSet(IconSetValues IconSetType)
        {
            switch (IconSetType)
            {
                case IconSetValues.FiveArrows:
                    this.SetCustomIconSet(new SLFiveIconSetOptions(SLFiveIconSetValues.FiveArrows));
                    break;
                case IconSetValues.FiveArrowsGray:
                    this.SetCustomIconSet(new SLFiveIconSetOptions(SLFiveIconSetValues.FiveArrowsGray));
                    break;
                case IconSetValues.FiveQuarters:
                    this.SetCustomIconSet(new SLFiveIconSetOptions(SLFiveIconSetValues.FiveQuarters));
                    break;
                case IconSetValues.FiveRating:
                    this.SetCustomIconSet(new SLFiveIconSetOptions(SLFiveIconSetValues.FiveRating));
                    break;
                case IconSetValues.FourArrows:
                    this.SetCustomIconSet(new SLFourIconSetOptions(SLFourIconSetValues.FourArrows));
                    break;
                case IconSetValues.FourArrowsGray:
                    this.SetCustomIconSet(new SLFourIconSetOptions(SLFourIconSetValues.FourArrowsGray));
                    break;
                case IconSetValues.FourRating:
                    this.SetCustomIconSet(new SLFourIconSetOptions(SLFourIconSetValues.FourRating));
                    break;
                case IconSetValues.FourRedToBlack:
                    this.SetCustomIconSet(new SLFourIconSetOptions(SLFourIconSetValues.FourRedToBlack));
                    break;
                case IconSetValues.FourTrafficLights:
                    this.SetCustomIconSet(new SLFourIconSetOptions(SLFourIconSetValues.FourTrafficLights));
                    break;
                case IconSetValues.ThreeArrows:
                    this.SetCustomIconSet(new SLThreeIconSetOptions(SLThreeIconSetValues.ThreeArrows));
                    break;
                case IconSetValues.ThreeArrowsGray:
                    this.SetCustomIconSet(new SLThreeIconSetOptions(SLThreeIconSetValues.ThreeArrowsGray));
                    break;
                case IconSetValues.ThreeFlags:
                    this.SetCustomIconSet(new SLThreeIconSetOptions(SLThreeIconSetValues.ThreeFlags));
                    break;
                case IconSetValues.ThreeSigns:
                    this.SetCustomIconSet(new SLThreeIconSetOptions(SLThreeIconSetValues.ThreeSigns));
                    break;
                case IconSetValues.ThreeSymbols:
                    this.SetCustomIconSet(new SLThreeIconSetOptions(SLThreeIconSetValues.ThreeSymbols));
                    break;
                case IconSetValues.ThreeSymbols2:
                    this.SetCustomIconSet(new SLThreeIconSetOptions(SLThreeIconSetValues.ThreeSymbols2));
                    break;
                case IconSetValues.ThreeTrafficLights1:
                    this.SetCustomIconSet(new SLThreeIconSetOptions(SLThreeIconSetValues.ThreeTrafficLights1));
                    break;
                case IconSetValues.ThreeTrafficLights2:
                    this.SetCustomIconSet(new SLThreeIconSetOptions(SLThreeIconSetValues.ThreeTrafficLights2));
                    break;
            }
        }

        /// <summary>
        /// Set a custom 3-icon set.
        /// </summary>
        /// <param name="IconSetType">The type of 3-icon set.</param>
        /// <param name="ReverseIconOrder">True to reverse the order of the icons. False to use the default order.</param>
        /// <param name="ShowIconOnly">True to show only icons. False to show both icon and value.</param>
        /// <param name="GreaterThanOrEqual2">True if values are to be greater than or equal to the 2nd range value. False if values are to be strictly greater than.</param>
        /// <param name="Value2">The 2nd range value.</param>
        /// <param name="Type2">The conditional format type for the 2nd range value.</param>
        /// <param name="GreaterThanOrEqual3">True if values are to be greater than or equal to the 3rd range value. False if values are to be strictly greater than.</param>
        /// <param name="Value3">The 3rd range value.</param>
        /// <param name="Type3">The conditional format type for the 3rd range value.</param>
        public void SetCustomIconSet(SLThreeIconSetValues IconSetType, bool ReverseIconOrder, bool ShowIconOnly,
            bool GreaterThanOrEqual2, string Value2, SLConditionalFormatRangeValues Type2,
            bool GreaterThanOrEqual3, string Value3, SLConditionalFormatRangeValues Type3)
        {
            SLThreeIconSetOptions Options = new SLThreeIconSetOptions(IconSetType);
            Options.ReverseIconOrder = ReverseIconOrder;
            Options.ShowIconOnly = ShowIconOnly;

            Options.GreaterThanOrEqual2 = GreaterThanOrEqual2;
            Options.Value2 = Value2;
            Options.Type2 = Type2;

            Options.GreaterThanOrEqual3 = GreaterThanOrEqual3;
            Options.Value3 = Value3;
            Options.Type3 = Type3;

            this.SetCustomIconSet(Options);
        }

        /// <summary>
        /// Set a custom 4-icon set.
        /// </summary>
        /// <param name="IconSetType">The type of 4-icon set.</param>
        /// <param name="ReverseIconOrder">True to reverse the order of the icons. False to use the default order.</param>
        /// <param name="ShowIconOnly">True to show only icons. False to show both icon and value.</param>
        /// <param name="GreaterThanOrEqual2">True if values are to be greater than or equal to the 2nd range value. False if values are to be strictly greater than.</param>
        /// <param name="Value2">The 2nd range value.</param>
        /// <param name="Type2">The conditional format type for the 2nd range value.</param>
        /// <param name="GreaterThanOrEqual3">True if values are to be greater than or equal to the 3rd range value. False if values are to be strictly greater than.</param>
        /// <param name="Value3">The 3rd range value.</param>
        /// <param name="Type3">The conditional format type for the 3rd range value.</param>
        /// <param name="GreaterThanOrEqual4">True if values are to be greater than or equal to the 4th range value. False if values are to be strictly greater than.</param>
        /// <param name="Value4">The 4th range value.</param>
        /// <param name="Type4">The conditional format type for the 4th range value.</param>
        public void SetCustomIconSet(SLFourIconSetValues IconSetType, bool ReverseIconOrder, bool ShowIconOnly,
            bool GreaterThanOrEqual2, string Value2, SLConditionalFormatRangeValues Type2,
            bool GreaterThanOrEqual3, string Value3, SLConditionalFormatRangeValues Type3,
            bool GreaterThanOrEqual4, string Value4, SLConditionalFormatRangeValues Type4)
        {
            SLFourIconSetOptions Options = new SLFourIconSetOptions(IconSetType);
            Options.ReverseIconOrder = ReverseIconOrder;
            Options.ShowIconOnly = ShowIconOnly;

            Options.GreaterThanOrEqual2 = GreaterThanOrEqual2;
            Options.Value2 = Value2;
            Options.Type2 = Type2;

            Options.GreaterThanOrEqual3 = GreaterThanOrEqual3;
            Options.Value3 = Value3;
            Options.Type3 = Type3;

            Options.GreaterThanOrEqual4 = GreaterThanOrEqual4;
            Options.Value4 = Value4;
            Options.Type4 = Type4;

            this.SetCustomIconSet(Options);
        }

        /// <summary>
        /// Set a custom 5-icon set.
        /// </summary>
        /// <param name="IconSetType">The type of 5-icon set.</param>
        /// <param name="ReverseIconOrder">True to reverse the order of the icons. False to use the default order.</param>
        /// <param name="ShowIconOnly">True to show only icons. False to show both icon and value.</param>
        /// <param name="GreaterThanOrEqual2">True if values are to be greater than or equal to the 2nd range value. False if values are to be strictly greater than.</param>
        /// <param name="Value2">The 2nd range value.</param>
        /// <param name="Type2">The conditional format type for the 2nd range value.</param>
        /// <param name="GreaterThanOrEqual3">True if values are to be greater than or equal to the 3rd range value. False if values are to be strictly greater than.</param>
        /// <param name="Value3">The 3rd range value.</param>
        /// <param name="Type3">The conditional format type for the 3rd range value.</param>
        /// <param name="GreaterThanOrEqual4">True if values are to be greater than or equal to the 4th range value. False if values are to be strictly greater than.</param>
        /// <param name="Value4">The 4th range value.</param>
        /// <param name="Type4">The conditional format type for the 4th range value.</param>
        /// <param name="GreaterThanOrEqual5">True if values are to be greater than or equal to the 5th range value. False if values are to be strictly greater than.</param>
        /// <param name="Value5">The 5th range value.</param>
        /// <param name="Type5">The conditional format type for the 5th range value.</param>
        public void SetCustomIconSet(SLFiveIconSetValues IconSetType, bool ReverseIconOrder, bool ShowIconOnly,
            bool GreaterThanOrEqual2, string Value2, SLConditionalFormatRangeValues Type2,
            bool GreaterThanOrEqual3, string Value3, SLConditionalFormatRangeValues Type3,
            bool GreaterThanOrEqual4, string Value4, SLConditionalFormatRangeValues Type4,
            bool GreaterThanOrEqual5, string Value5, SLConditionalFormatRangeValues Type5)
        {
            SLFiveIconSetOptions Options = new SLFiveIconSetOptions(IconSetType);
            Options.ReverseIconOrder = ReverseIconOrder;
            Options.ShowIconOnly = ShowIconOnly;

            Options.GreaterThanOrEqual2 = GreaterThanOrEqual2;
            Options.Value2 = Value2;
            Options.Type2 = Type2;

            Options.GreaterThanOrEqual3 = GreaterThanOrEqual3;
            Options.Value3 = Value3;
            Options.Type3 = Type3;

            Options.GreaterThanOrEqual4 = GreaterThanOrEqual4;
            Options.Value4 = Value4;
            Options.Type4 = Type4;

            Options.GreaterThanOrEqual5 = GreaterThanOrEqual5;
            Options.Value5 = Value5;
            Options.Type5 = Type5;

            this.SetCustomIconSet(Options);
        }

        /// <summary>
        /// Set a custom 3-icon set.
        /// </summary>
        /// <param name="Options">3-icon set options.</param>
        public void SetCustomIconSet(SLThreeIconSetOptions Options)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.IconSet;
            cfr.IconSet.Reverse = Options.ReverseIconOrder;
            cfr.IconSet.ShowValue = !Options.ShowIconOnly;

            switch (Options.IconSetType)
            {
                case SLThreeIconSetValues.ThreeArrows:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeArrows;
                    break;
                case SLThreeIconSetValues.ThreeArrowsGray:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeArrowsGray;
                    break;
                case SLThreeIconSetValues.ThreeFlags:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeFlags;
                    break;
                case SLThreeIconSetValues.ThreeSigns:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeSigns;
                    break;
                case SLThreeIconSetValues.ThreeStars:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeStars;
                    break;
                case SLThreeIconSetValues.ThreeSymbols:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeSymbols;
                    break;
                case SLThreeIconSetValues.ThreeSymbols2:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeSymbols2;
                    break;
                case SLThreeIconSetValues.ThreeTrafficLights1:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeTrafficLights1;
                    break;
                case SLThreeIconSetValues.ThreeTrafficLights2:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeTrafficLights2;
                    break;
                case SLThreeIconSetValues.ThreeTriangles:
                    cfr.IconSet.IconSetType = SLIconSetValues.ThreeTriangles;
                    break;
            }

            cfr.IconSet.Is2010 = SLIconSet.Is2010IconSet(cfr.IconSet.IconSetType);

            if (Options.IsCustomIcon)
            {
                X14.IconSetTypeValues istv = X14.IconSetTypeValues.ThreeTrafficLights1;
                uint iIconId = 0;
                
                SLIconSet.TranslateCustomIcon(Options.Icon1, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                SLIconSet.TranslateCustomIcon(Options.Icon2, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                SLIconSet.TranslateCustomIcon(Options.Icon3, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                cfr.IconSet.Is2010 = true;
            }

            SLConditionalFormatValueObject cfvo;

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = ConditionalFormatValueObjectValues.Percent;
            cfvo.Val = "0";
            cfr.IconSet.Cfvos.Add(cfvo);

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.TranslateRangeValues(Options.Type2);
            cfvo.Val = Options.Value2;
            cfvo.GreaterThanOrEqual = Options.GreaterThanOrEqual2;
            cfr.IconSet.Cfvos.Add(cfvo);

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.TranslateRangeValues(Options.Type3);
            cfvo.Val = Options.Value3;
            cfvo.GreaterThanOrEqual = Options.GreaterThanOrEqual3;
            cfr.IconSet.Cfvos.Add(cfvo);

            cfr.HasIconSet = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Set a custom 4-icon set.
        /// </summary>
        /// <param name="Options">4-icon set options.</param>
        public void SetCustomIconSet(SLFourIconSetOptions Options)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.IconSet;
            cfr.IconSet.Reverse = Options.ReverseIconOrder;
            cfr.IconSet.ShowValue = !Options.ShowIconOnly;

            switch (Options.IconSetType)
            {
                case SLFourIconSetValues.FourArrows:
                    cfr.IconSet.IconSetType = SLIconSetValues.FourArrows;
                    break;
                case SLFourIconSetValues.FourArrowsGray:
                    cfr.IconSet.IconSetType = SLIconSetValues.FourArrowsGray;
                    break;
                case SLFourIconSetValues.FourRating:
                    cfr.IconSet.IconSetType = SLIconSetValues.FourRating;
                    break;
                case SLFourIconSetValues.FourRedToBlack:
                    cfr.IconSet.IconSetType = SLIconSetValues.FourRedToBlack;
                    break;
                case SLFourIconSetValues.FourTrafficLights:
                    cfr.IconSet.IconSetType = SLIconSetValues.FourTrafficLights;
                    break;
            }

            cfr.IconSet.Is2010 = SLIconSet.Is2010IconSet(cfr.IconSet.IconSetType);

            if (Options.IsCustomIcon)
            {
                X14.IconSetTypeValues istv = X14.IconSetTypeValues.ThreeTrafficLights1;
                uint iIconId = 0;

                SLIconSet.TranslateCustomIcon(Options.Icon1, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                SLIconSet.TranslateCustomIcon(Options.Icon2, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                SLIconSet.TranslateCustomIcon(Options.Icon3, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                SLIconSet.TranslateCustomIcon(Options.Icon4, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                cfr.IconSet.Is2010 = true;
            }

            SLConditionalFormatValueObject cfvo;

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = ConditionalFormatValueObjectValues.Percent;
            cfvo.Val = "0";
            cfr.IconSet.Cfvos.Add(cfvo);

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.TranslateRangeValues(Options.Type2);
            cfvo.Val = Options.Value2;
            cfvo.GreaterThanOrEqual = Options.GreaterThanOrEqual2;
            cfr.IconSet.Cfvos.Add(cfvo);

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.TranslateRangeValues(Options.Type3);
            cfvo.Val = Options.Value3;
            cfvo.GreaterThanOrEqual = Options.GreaterThanOrEqual3;
            cfr.IconSet.Cfvos.Add(cfvo);

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.TranslateRangeValues(Options.Type4);
            cfvo.Val = Options.Value4;
            cfvo.GreaterThanOrEqual = Options.GreaterThanOrEqual4;
            cfr.IconSet.Cfvos.Add(cfvo);

            cfr.HasIconSet = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Set a custom 5-icon set.
        /// </summary>
        /// <param name="Options">5-icon set options.</param>
        public void SetCustomIconSet(SLFiveIconSetOptions Options)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.IconSet;
            cfr.IconSet.Reverse = Options.ReverseIconOrder;
            cfr.IconSet.ShowValue = !Options.ShowIconOnly;

            switch (Options.IconSetType)
            {
                case SLFiveIconSetValues.FiveArrows:
                    cfr.IconSet.IconSetType = SLIconSetValues.FiveArrows;
                    break;
                case SLFiveIconSetValues.FiveArrowsGray:
                    cfr.IconSet.IconSetType = SLIconSetValues.FiveArrowsGray;
                    break;
                case SLFiveIconSetValues.FiveBoxes:
                    cfr.IconSet.IconSetType = SLIconSetValues.FiveBoxes;
                    break;
                case SLFiveIconSetValues.FiveQuarters:
                    cfr.IconSet.IconSetType = SLIconSetValues.FiveQuarters;
                    break;
                case SLFiveIconSetValues.FiveRating:
                    cfr.IconSet.IconSetType = SLIconSetValues.FiveRating;
                    break;
            }

            cfr.IconSet.Is2010 = SLIconSet.Is2010IconSet(cfr.IconSet.IconSetType);

            if (Options.IsCustomIcon)
            {
                X14.IconSetTypeValues istv = X14.IconSetTypeValues.ThreeTrafficLights1;
                uint iIconId = 0;

                SLIconSet.TranslateCustomIcon(Options.Icon1, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                SLIconSet.TranslateCustomIcon(Options.Icon2, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                SLIconSet.TranslateCustomIcon(Options.Icon3, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                SLIconSet.TranslateCustomIcon(Options.Icon4, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                SLIconSet.TranslateCustomIcon(Options.Icon5, out istv, out iIconId);
                cfr.IconSet.CustomIcons.Add(new SLConditionalFormattingIcon2010() { IconSet = istv, IconId = iIconId });

                cfr.IconSet.Is2010 = true;
            }

            SLConditionalFormatValueObject cfvo;

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = ConditionalFormatValueObjectValues.Percent;
            cfvo.Val = "0";
            cfr.IconSet.Cfvos.Add(cfvo);

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.TranslateRangeValues(Options.Type2);
            cfvo.Val = Options.Value2;
            cfvo.GreaterThanOrEqual = Options.GreaterThanOrEqual2;
            cfr.IconSet.Cfvos.Add(cfvo);

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.TranslateRangeValues(Options.Type3);
            cfvo.Val = Options.Value3;
            cfvo.GreaterThanOrEqual = Options.GreaterThanOrEqual3;
            cfr.IconSet.Cfvos.Add(cfvo);

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.TranslateRangeValues(Options.Type4);
            cfvo.Val = Options.Value4;
            cfvo.GreaterThanOrEqual = Options.GreaterThanOrEqual4;
            cfr.IconSet.Cfvos.Add(cfvo);

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.TranslateRangeValues(Options.Type5);
            cfvo.Val = Options.Value5;
            cfvo.GreaterThanOrEqual = Options.GreaterThanOrEqual5;
            cfr.IconSet.Cfvos.Add(cfvo);

            cfr.HasIconSet = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values greater than a given value.
        /// </summary>
        /// <param name="IncludeEquality">True for greater than or equal to. False for strictly greater than.</param>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsGreaterThan(bool IncludeEquality, string Value, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsGreaterThan(IncludeEquality, Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values greater than a given value.
        /// </summary>
        /// <param name="IncludeEquality">True for greater than or equal to. False for strictly greater than.</param>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsGreaterThan(bool IncludeEquality, string Value, SLStyle HighlightStyle)
        {
            this.HighlightCellsGreaterThan(IncludeEquality, Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsGreaterThan(bool IncludeEquality, string Value, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.CellIs;
            cfr.Operator = IncludeEquality ? ConditionalFormattingOperatorValues.GreaterThanOrEqual : ConditionalFormattingOperatorValues.GreaterThan;
            cfr.HasOperator = true;

            // 30 Jan 2017: non-numeric values used to be automatically quoted
            // but might be a formula expression like $E$1
            if (Value.StartsWith("="))
            {
                string s = Value;
                if (Value.Length > 1)
                {
                    s = Value.Substring(1);
                }
                cfr.Formulas.Add(new Formula(s));
            }
            else
            {
                cfr.Formulas.Add(this.GetFormulaFromText(Value));
            }

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values less than a given value.
        /// </summary>
        /// <param name="IncludeEquality">True for less than or equal to. False for strictly less than.</param>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsLessThan(bool IncludeEquality, string Value, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsLessThan(IncludeEquality, Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values less than a given value.
        /// </summary>
        /// <param name="IncludeEquality">True for less than or equal to. False for strictly less than.</param>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsLessThan(bool IncludeEquality, string Value, SLStyle HighlightStyle)
        {
            this.HighlightCellsLessThan(IncludeEquality, Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsLessThan(bool IncludeEquality, string Value, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.CellIs;
            cfr.Operator = IncludeEquality ? ConditionalFormattingOperatorValues.LessThanOrEqual : ConditionalFormattingOperatorValues.LessThan;
            cfr.HasOperator = true;

            // 30 Jan 2017: non-numeric values used to be automatically quoted
            // but might be a formula expression like $E$1
            if (Value.StartsWith("="))
            {
                string s = Value;
                if (Value.Length > 1)
                {
                    s = Value.Substring(1);
                }
                cfr.Formulas.Add(new Formula(s));
            }
            else
            {
                cfr.Formulas.Add(this.GetFormulaFromText(Value));
            }

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values between 2 given values.
        /// </summary>
        /// <param name="IsBetween">True for between the 2 given values. False for not between the 2 given values.</param>
        /// <param name="Value1">The 1st value to be compared with.</param>
        /// <param name="Value2">The 2nd value to be compared with.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsBetween(bool IsBetween, string Value1, string Value2, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsBetween(IsBetween, Value1, Value2, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values between 2 given values.
        /// </summary>
        /// <param name="IsBetween">True for between the 2 given values. False for not between the 2 given values.</param>
        /// <param name="Value1">The 1st value to be compared with.</param>
        /// <param name="Value2">The 2nd value to be compared with.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsBetween(bool IsBetween, string Value1, string Value2, SLStyle HighlightStyle)
        {
            this.HighlightCellsBetween(IsBetween, Value1, Value2, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsBetween(bool IsBetween, string Value1, string Value2, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.CellIs;
            cfr.Operator = IsBetween ? ConditionalFormattingOperatorValues.Between : ConditionalFormattingOperatorValues.NotBetween;
            cfr.HasOperator = true;

            string s;

            // 30 Jan 2017: non-numeric values used to be automatically quoted
            // but might be a formula expression like $E$1
            if (Value1.StartsWith("="))
            {
                s = Value1;
                if (Value1.Length > 1)
                {
                    s = Value1.Substring(1);
                }
                cfr.Formulas.Add(new Formula(s));
            }
            else
            {
                cfr.Formulas.Add(this.GetFormulaFromText(Value1));
            }

            if (Value2.StartsWith("="))
            {
                s = Value2;
                if (Value2.Length > 1)
                {
                    s = Value2.Substring(1);
                }
                cfr.Formulas.Add(new Formula(s));
            }
            else
            {
                cfr.Formulas.Add(this.GetFormulaFromText(Value2));
            }

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values equal to a given value.
        /// </summary>
        /// <param name="IsEqual">True for equal to given value. False for not equal to given value.</param>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsEqual(bool IsEqual, string Value, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsEqual(IsEqual, Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values equal to a given value.
        /// </summary>
        /// <param name="IsEqual">True for equal to given value. False for not equal to given value.</param>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsEqual(bool IsEqual, string Value, SLStyle HighlightStyle)
        {
            this.HighlightCellsEqual(IsEqual, Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsEqual(bool IsEqual, string Value, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.CellIs;
            cfr.Operator = IsEqual ? ConditionalFormattingOperatorValues.Equal : ConditionalFormattingOperatorValues.NotEqual;
            cfr.HasOperator = true;

            // 30 Jan 2017: non-numeric values used to be automatically quoted
            // but might be a formula expression like $E$1
            if (Value.StartsWith("="))
            {
                string s = Value;
                if (Value.Length > 1)
                {
                    s = Value.Substring(1);
                }
                cfr.Formulas.Add(new Formula(s));
            }
            else
            {
                cfr.Formulas.Add(this.GetFormulaFromText(Value));
            }

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values containing a given value.
        /// </summary>
        /// <param name="IsContaining">True for containing given value. False for not containing given value.</param>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsContainingText(bool IsContaining, string Value, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsContainingText(IsContaining, Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values containing a given value.
        /// </summary>
        /// <param name="IsContaining">True for containing given value. False for not containing given value.</param>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsContainingText(bool IsContaining, string Value, SLStyle HighlightStyle)
        {
            this.HighlightCellsContainingText(IsContaining, Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsContainingText(bool IsContaining, string Value, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();

            string s = Value;
            bool bIsFormula = false;
            // 30 Jan 2017: non-numeric values used to be automatically quoted
            // but might be a formula expression like $E$1
            if (Value.StartsWith("="))
            {
                bIsFormula = true;
                if (Value.Length > 1)
                {
                    s = Value.Substring(1);
                }
            }

            cfr.Text = s;

            Formula f = new Formula();
            string sRef = string.Empty;
            if (this.SequenceOfReferences.Count > 0)
            {
                sRef = SLTool.ToCellReference(this.SequenceOfReferences[0].StartRowIndex, this.SequenceOfReferences[0].StartColumnIndex);
            }
            if (IsContaining)
            {
                cfr.Type = ConditionalFormatValues.ContainsText;
                cfr.Operator = ConditionalFormattingOperatorValues.ContainsText;
                cfr.HasOperator = true;
                f.Text = string.Format("NOT(ISERROR(SEARCH({0},{1})))", bIsFormula ? s : this.GetCleanedStringFromText(Value), sRef);
            }
            else
            {
                cfr.Type = ConditionalFormatValues.NotContainsText;
                cfr.Operator = ConditionalFormattingOperatorValues.NotContains;
                cfr.HasOperator = true;
                f.Text = string.Format("ISERROR(SEARCH({0},{1}))", bIsFormula ? s : this.GetCleanedStringFromText(Value), sRef);
            }
            cfr.Formulas.Add(f);

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values beginning with a given value.
        /// </summary>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsBeginningWith(string Value, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsBeginningWith(Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values beginning with a given value.
        /// </summary>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsBeginningWith(string Value, SLStyle HighlightStyle)
        {
            this.HighlightCellsBeginningWith(Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsBeginningWith(string Value, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();

            string s = Value;
            bool bIsFormula = false;
            // 30 Jan 2017: non-numeric values used to be automatically quoted
            // but might be a formula expression like $E$1
            if (Value.StartsWith("="))
            {
                bIsFormula = true;
                if (Value.Length > 1)
                {
                    s = Value.Substring(1);
                }
            }

            cfr.Text = s;

            Formula f = new Formula();
            string sRef = string.Empty;
            if (this.SequenceOfReferences.Count > 0)
            {
                sRef = SLTool.ToCellReference(this.SequenceOfReferences[0].StartRowIndex, this.SequenceOfReferences[0].StartColumnIndex);
            }
            cfr.Type = ConditionalFormatValues.BeginsWith;
            cfr.Operator = ConditionalFormattingOperatorValues.BeginsWith;
            cfr.HasOperator = true;
            f.Text = string.Format("LEFT({0},{1})={2}", sRef, bIsFormula ? string.Format("LEN({0})", s) : Value.Length.ToString(), bIsFormula ? s : this.GetCleanedStringFromText(Value));
            cfr.Formulas.Add(f);

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values ending with a given value.
        /// </summary>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsEndingWith(string Value, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsEndingWith(Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values ending with a given value.
        /// </summary>
        /// <param name="Value">The value to be compared with.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsEndingWith(string Value, SLStyle HighlightStyle)
        {
            this.HighlightCellsEndingWith(Value, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsEndingWith(string Value, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();

            string s = Value;
            bool bIsFormula = false;
            // 30 Jan 2017: non-numeric values used to be automatically quoted
            // but might be a formula expression like $E$1
            if (Value.StartsWith("="))
            {
                bIsFormula = true;
                if (Value.Length > 1)
                {
                    s = Value.Substring(1);
                }
            }

            cfr.Text = s;

            Formula f = new Formula();
            string sRef = string.Empty;
            if (this.SequenceOfReferences.Count > 0)
            {
                sRef = SLTool.ToCellReference(this.SequenceOfReferences[0].StartRowIndex, this.SequenceOfReferences[0].StartColumnIndex);
            }
            cfr.Type = ConditionalFormatValues.EndsWith;
            cfr.Operator = ConditionalFormattingOperatorValues.EndsWith;
            cfr.HasOperator = true;
            f.Text = string.Format("RIGHT({0},{1})={2}", sRef, bIsFormula ? string.Format("LEN({0})", s) : Value.Length.ToString(), bIsFormula ? s : this.GetCleanedStringFromText(Value));
            cfr.Formulas.Add(f);

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells that are blank.
        /// </summary>
        /// <param name="ContainsBlanks">True for containing blanks. False for not containing blanks.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsContainingBlanks(bool ContainsBlanks, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsContainingBlanks(ContainsBlanks, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells that are blank.
        /// </summary>
        /// <param name="ContainsBlanks">True for containing blanks. False for not containing blanks.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsContainingBlanks(bool ContainsBlanks, SLStyle HighlightStyle)
        {
            this.HighlightCellsContainingBlanks(ContainsBlanks, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsContainingBlanks(bool ContainsBlanks, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();

            Formula f = new Formula();
            string sRef = string.Empty;
            if (this.SequenceOfReferences.Count > 0)
            {
                sRef = SLTool.ToCellReference(this.SequenceOfReferences[0].StartRowIndex, this.SequenceOfReferences[0].StartColumnIndex);
            }
            if (ContainsBlanks)
            {
                cfr.Type = ConditionalFormatValues.ContainsBlanks;
                f.Text = string.Format("LEN(TRIM({0}))=0", sRef);
            }
            else
            {
                cfr.Type = ConditionalFormatValues.NotContainsBlanks;
                f.Text = string.Format("LEN(TRIM({0}))>0", sRef);
            }
            cfr.Formulas.Add(f);

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells containing errors.
        /// </summary>
        /// <param name="ContainsErrors">True for containing errors. False for not containing errors.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsContainingErrors(bool ContainsErrors, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsContainingErrors(ContainsErrors, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells containing errors.
        /// </summary>
        /// <param name="ContainsErrors">True for containing errors. False for not containing errors.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsContainingErrors(bool ContainsErrors, SLStyle HighlightStyle)
        {
            this.HighlightCellsContainingErrors(ContainsErrors, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsContainingErrors(bool ContainsErrors, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();

            Formula f = new Formula();
            string sRef = string.Empty;
            if (this.SequenceOfReferences.Count > 0)
            {
                sRef = SLTool.ToCellReference(this.SequenceOfReferences[0].StartRowIndex, this.SequenceOfReferences[0].StartColumnIndex);
            }
            if (ContainsErrors)
            {
                cfr.Type = ConditionalFormatValues.ContainsErrors;
                f.Text = string.Format("ISERROR({0})", sRef);
            }
            else
            {
                cfr.Type = ConditionalFormatValues.NotContainsErrors;
                f.Text = string.Format("NOT(ISERROR({0}))", sRef);
            }
            cfr.Formulas.Add(f);

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with date values occurring according to a given time period.
        /// </summary>
        /// <param name="DatePeriod">A given time period.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsWithDatesOccurring(TimePeriodValues DatePeriod, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsWithDatesOccurring(DatePeriod, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with date values occurring according to a given time period.
        /// </summary>
        /// <param name="DatePeriod">A given time period.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsWithDatesOccurring(TimePeriodValues DatePeriod, SLStyle HighlightStyle)
        {
            this.HighlightCellsWithDatesOccurring(DatePeriod, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsWithDatesOccurring(TimePeriodValues DatePeriod, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.TimePeriod;
            cfr.TimePeriod = DatePeriod;
            cfr.HasTimePeriod = true;

            Formula f = new Formula();
            string sRef = string.Empty;
            if (this.SequenceOfReferences.Count > 0)
            {
                sRef = SLTool.ToCellReference(this.SequenceOfReferences[0].StartRowIndex, this.SequenceOfReferences[0].StartColumnIndex);
            }
            switch (DatePeriod)
            {
                case TimePeriodValues.Yesterday:
                    f.Text = string.Format("FLOOR({0},1)=TODAY()-1", sRef);
                    break;
                case TimePeriodValues.Today:
                    f.Text = string.Format("FLOOR({0},1)=TODAY()", sRef);
                    break;
                case TimePeriodValues.Tomorrow:
                    f.Text = string.Format("FLOOR({0},1)=TODAY()+1", sRef);
                    break;
                case TimePeriodValues.Last7Days:
                    f.Text = string.Format("AND(TODAY()-FLOOR({0},1)<=6,FLOOR({0},1)<=TODAY())", sRef);
                    break;
                case TimePeriodValues.LastWeek:
                    f.Text = string.Format("AND(TODAY()-ROUNDDOWN({0},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({0},0)<(WEEKDAY(TODAY())+7))", sRef);
                    break;
                case TimePeriodValues.ThisWeek:
                    f.Text = string.Format("AND(TODAY()-ROUNDDOWN({0},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({0},0)-TODAY()<=7-WEEKDAY(TODAY()))", sRef);
                    break;
                case TimePeriodValues.NextWeek:
                    f.Text = string.Format("AND(ROUNDDOWN({0},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({0},0)-TODAY()<(15-WEEKDAY(TODAY())))", sRef);
                    break;
                case TimePeriodValues.LastMonth:
                    f.Text = string.Format("AND(MONTH({0})=MONTH(EDATE(TODAY(),0-1)),YEAR({0})=YEAR(EDATE(TODAY(),0-1)))", sRef);
                    break;
                case TimePeriodValues.ThisMonth:
                    f.Text = string.Format("AND(MONTH({0})=MONTH(TODAY()),YEAR({0})=YEAR(TODAY()))", sRef);
                    break;
                case TimePeriodValues.NextMonth:
                    f.Text = string.Format("AND(MONTH({0})=MONTH(TODAY())+1,OR(YEAR({0})=YEAR(TODAY()),AND(MONTH({0})=12,YEAR({0})=YEAR(TODAY())+1)))", sRef);
                    break;
            }
            cfr.Formulas.Add(f);

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with duplicate values.
        /// </summary>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsWithDuplicates(SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsWithDuplicates(this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with duplicate values.
        /// </summary>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsWithDuplicates(SLStyle HighlightStyle)
        {
            this.HighlightCellsWithDuplicates(this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsWithDuplicates(SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.DuplicateValues;

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with unique values.
        /// </summary>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsWithUniques(SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsWithUniques(this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with unique values.
        /// </summary>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsWithUniques(SLStyle HighlightStyle)
        {
            this.HighlightCellsWithUniques(this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsWithUniques(SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.UniqueValues;

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values in the top/bottom range.
        /// </summary>
        /// <param name="IsTopRange">True if in the top range. False if in the bottom range.</param>
        /// <param name="Rank">The value of X in "Top/Bottom X". If <paramref name="IsPercent"/> is true, then X refers to X%, otherwise it's X number of items.</param>
        /// <param name="IsPercent">True if referring to percentage. False if referring to number of items.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsInTopRange(bool IsTopRange, uint Rank, bool IsPercent, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsInTopRange(IsTopRange, Rank, IsPercent, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values in the top/bottom range.
        /// </summary>
        /// <param name="IsTopRange">True if in the top range. False if in the bottom range.</param>
        /// <param name="Rank">The value of X in "Top/Bottom X". If <paramref name="IsPercent"/> is true, then X refers to X%, otherwise it's X number of items.</param>
        /// <param name="IsPercent">True if referring to percentage. False if referring to number of items.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsInTopRange(bool IsTopRange, uint Rank, bool IsPercent, SLStyle HighlightStyle)
        {
            this.HighlightCellsInTopRange(IsTopRange, Rank, IsPercent, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsInTopRange(bool IsTopRange, uint Rank, bool IsPercent, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.Top10;
            cfr.Bottom = !IsTopRange;
            cfr.Rank = Rank;
            cfr.Percent = IsPercent;

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values compared to the average.
        /// </summary>
        /// <param name="AverageType">The type of comparison to the average.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsAboveAverage(SLHighlightCellsAboveAverageValues AverageType, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsAboveAverage(AverageType, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values compared to the average.
        /// </summary>
        /// <param name="AverageType">The type of comparison to the average.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsAboveAverage(SLHighlightCellsAboveAverageValues AverageType, SLStyle HighlightStyle)
        {
            this.HighlightCellsAboveAverage(AverageType, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsAboveAverage(SLHighlightCellsAboveAverageValues AverageType, SLDifferentialFormat HighlightStyle)
        {
            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.AboveAverage;

            switch (AverageType)
            {
                case SLHighlightCellsAboveAverageValues.Above:
                    // that's all is needed!
                    break;
                case SLHighlightCellsAboveAverageValues.Below:
                    cfr.AboveAverage = false;
                    break;
                case SLHighlightCellsAboveAverageValues.EqualOrAbove:
                    cfr.EqualAverage = true;
                    break;
                case SLHighlightCellsAboveAverageValues.EqualOrBelow:
                    cfr.EqualAverage = true;
                    cfr.AboveAverage = false;
                    break;
                case SLHighlightCellsAboveAverageValues.OneStdDevAbove:
                    cfr.StdDev = 1;
                    break;
                case SLHighlightCellsAboveAverageValues.OneStdDevBelow:
                    cfr.AboveAverage = false;
                    cfr.StdDev = 1;
                    break;
                case SLHighlightCellsAboveAverageValues.TwoStdDevAbove:
                    cfr.StdDev = 2;
                    break;
                case SLHighlightCellsAboveAverageValues.TwoStdDevBelow:
                    cfr.AboveAverage = false;
                    cfr.StdDev = 2;
                    break;
                case SLHighlightCellsAboveAverageValues.ThreeStdDevAbove:
                    cfr.StdDev = 3;
                    break;
                case SLHighlightCellsAboveAverageValues.ThreeStdDevBelow:
                    cfr.AboveAverage = false;
                    cfr.StdDev = 3;
                    break;
            }

            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        /// <summary>
        /// Highlight cells with values according to a formula.
        /// </summary>
        /// <param name="Formula">The formula to apply.</param>
        /// <param name="HighlightStyle">A built-in highlight style.</param>
        public void HighlightCellsWithFormula(string Formula, SLHighlightCellsStyleValues HighlightStyle)
        {
            this.HighlightCellsWithFormula(Formula, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        /// <summary>
        /// Highlight cells with values according to a formula.
        /// </summary>
        /// <param name="Formula">The formula to apply.</param>
        /// <param name="HighlightStyle">A custom formatted style. Note that only number formats, fonts, borders and fills are used. Note further that for fonts, only italic/bold, underline, color and strikethrough settings are used.</param>
        public void HighlightCellsWithFormula(string Formula, SLStyle HighlightStyle)
        {
            this.HighlightCellsWithFormula(Formula, this.TranslateToDifferentialFormat(HighlightStyle));
        }

        private void HighlightCellsWithFormula(string Formula, SLDifferentialFormat HighlightStyle)
        {
            if (Formula.StartsWith("=")) Formula = Formula.Substring(1);

            SLConditionalFormattingRule cfr = new SLConditionalFormattingRule();
            cfr.Type = ConditionalFormatValues.Expression;
            cfr.Formulas.Add(new DocumentFormat.OpenXml.Spreadsheet.Formula(Formula));
            cfr.DifferentialFormat = HighlightStyle.Clone();
            cfr.HasDifferentialFormat = true;

            this.AppendRule(cfr);
        }

        internal Formula GetFormulaFromText(string Text)
        {
            Formula f = new Formula();
            double fTemp = 0.0;
            if (double.TryParse(Text, out fTemp))
            {
                f.Text = Text;
            }
            else
            {
                // double quotes are doubled
                Text = Text.Replace("\"", "\"\"");
                if (SLTool.ToPreserveSpace(Text))
                {
                    f.Space = SpaceProcessingModeValues.Preserve;
                }
                // double quotes are placed at the ends of the given value
                f.Text = string.Format("\"{0}\"", Text);
            }
            return f;
        }

        internal string GetCleanedStringFromText(string Text)
        {
            // double quotes are doubled
            Text = Text.Replace("\"", "\"\"");
            // double quotes are placed at the ends of the given value
            Text = string.Format("\"{0}\"", Text);
            return Text;
        }

        internal ConditionalFormatValueObjectValues TranslateRangeValues(SLConditionalFormatRangeValues RangeValue)
        {
            ConditionalFormatValueObjectValues cfvov = ConditionalFormatValueObjectValues.Number;
            switch (RangeValue)
            {
                case SLConditionalFormatRangeValues.Number:
                    cfvov = ConditionalFormatValueObjectValues.Number;
                    break;
                case SLConditionalFormatRangeValues.Percent:
                    cfvov = ConditionalFormatValueObjectValues.Percent;
                    break;
                case SLConditionalFormatRangeValues.Formula:
                    cfvov = ConditionalFormatValueObjectValues.Formula;
                    break;
                case SLConditionalFormatRangeValues.Percentile:
                    cfvov = ConditionalFormatValueObjectValues.Percentile;
                    break;
            }
            return cfvov;
        }

        internal SLConditionalFormatAutoMinMaxValues TranslateMinMaxValues(SLConditionalFormatMinMaxValues MinMaxValue)
        {
            SLConditionalFormatAutoMinMaxValues result = SLConditionalFormatAutoMinMaxValues.Value;
            switch (MinMaxValue)
            {
                case SLConditionalFormatMinMaxValues.Formula:
                    result = SLConditionalFormatAutoMinMaxValues.Formula;
                    break;
                case SLConditionalFormatMinMaxValues.Number:
                    result = SLConditionalFormatAutoMinMaxValues.Number;
                    break;
                case SLConditionalFormatMinMaxValues.Percent:
                    result = SLConditionalFormatAutoMinMaxValues.Percent;
                    break;
                case SLConditionalFormatMinMaxValues.Percentile:
                    result = SLConditionalFormatAutoMinMaxValues.Percentile;
                    break;
                case SLConditionalFormatMinMaxValues.Value:
                    result = SLConditionalFormatAutoMinMaxValues.Value;
                    break;
            }

            return result;
        }

        internal SLDifferentialFormat TranslateToDifferentialFormat(SLHighlightCellsStyleValues style)
        {
            SLDifferentialFormat df = new SLDifferentialFormat();
            switch (style)
            {
                case SLHighlightCellsStyleValues.LightRedFillWithDarkRedText:
                    df.Font.Condense = false;
                    df.Font.Extend = false;
                    df.Font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0x9C, 0x00, 0x06);
                    df.Fill.SetPatternBackgroundColor(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xC7, 0xCE));
                    break;
                case SLHighlightCellsStyleValues.YellowFillWithDarkYellowText:
                    df.Font.Condense = false;
                    df.Font.Extend = false;
                    df.Font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0x9C, 0x65, 0x00);
                    df.Fill.SetPatternBackgroundColor(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xEB, 0x9C));
                    break;
                case SLHighlightCellsStyleValues.GreenFillWithDarkGreenText:
                    df.Font.Condense = false;
                    df.Font.Extend = false;
                    df.Font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0x00, 0x61, 0x00);
                    df.Fill.SetPatternBackgroundColor(System.Drawing.Color.FromArgb(0xFF, 0xC6, 0xEF, 0xCE));
                    break;
                case SLHighlightCellsStyleValues.LightRedFill:
                    df.Fill.SetPatternBackgroundColor(System.Drawing.Color.FromArgb(0xFF, 0xFF, 0xC7, 0xCE));
                    break;
                case SLHighlightCellsStyleValues.RedText:
                    df.Font.Condense = false;
                    df.Font.Extend = false;
                    df.Font.FontColor = System.Drawing.Color.FromArgb(0xFF, 0x9C, 0x00, 0x06);
                    break;
                case SLHighlightCellsStyleValues.RedBorder:
                    df.Border.SetLeftBorder(BorderStyleValues.Thin, System.Drawing.Color.FromArgb(0xFF, 0x9C, 0x00, 0x06));
                    df.Border.SetRightBorder(BorderStyleValues.Thin, System.Drawing.Color.FromArgb(0xFF, 0x9C, 0x00, 0x06));
                    df.Border.SetTopBorder(BorderStyleValues.Thin, System.Drawing.Color.FromArgb(0xFF, 0x9C, 0x00, 0x06));
                    df.Border.SetBottomBorder(BorderStyleValues.Thin, System.Drawing.Color.FromArgb(0xFF, 0x9C, 0x00, 0x06));
                    break;
            }

            df.Sync();
            return df;
        }

        internal SLDifferentialFormat TranslateToDifferentialFormat(SLStyle style)
        {
            style.Sync();
            SLDifferentialFormat df = new SLDifferentialFormat();
            if (style.HasNumberingFormat) df.FormatCode = style.FormatCode;

            if (style.Font.Italic != null && style.Font.Italic.Value) df.Font.Italic = true;
            if (style.Font.Bold != null && style.Font.Bold.Value) df.Font.Bold = true;
            if (style.Font.HasUnderline) df.Font.Underline = style.Font.Underline;
            if (style.Font.HasFontColor)
            {
                df.Font.clrFontColor = style.Font.clrFontColor.Clone();
                df.Font.HasFontColor = true;
            }
            if (style.Font.Strike != null && style.Font.Strike.Value) df.Font.Strike = true;

            if (style.HasBorder) df.Border = style.Border.Clone();
            if (style.HasFill) df.Fill = style.Fill.Clone();

            df.Sync();
            return df;
        }

        internal void FromConditionalFormatting(ConditionalFormatting cf)
        {
            this.SetAllNull();

            if (cf.Pivot != null) this.Pivot = cf.Pivot.Value;

            if (cf.SequenceOfReferences != null)
            {
                this.SequenceOfReferences = SLTool.TranslateSeqRefToCellPointRange(cf.SequenceOfReferences);
            }

            using (OpenXmlReader oxr = OpenXmlReader.Create(cf))
            {
                SLConditionalFormattingRule cfr;
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(ConditionalFormattingRule))
                    {
                        cfr = new SLConditionalFormattingRule();
                        cfr.FromConditionalFormattingRule((ConditionalFormattingRule)oxr.LoadCurrentElement());
                        this.Rules.Add(cfr);
                    }
                }
            }
        }

        internal ConditionalFormatting ToConditionalFormatting()
        {
            ConditionalFormatting cf = new ConditionalFormatting();
            if (this.Pivot) cf.Pivot = this.Pivot;
            cf.SequenceOfReferences = SLTool.TranslateCellPointRangeToSeqRef(this.SequenceOfReferences);

            foreach (SLConditionalFormattingRule cfr in this.Rules)
            {
                cf.Append(cfr.ToConditionalFormattingRule());
            }

            return cf;
        }

        internal SLConditionalFormatting Clone()
        {
            SLConditionalFormatting cf = new SLConditionalFormatting();

            int i;
            cf.Rules = new List<SLConditionalFormattingRule>();
            for (i = 0; i < this.Rules.Count; ++i)
            {
                cf.Rules.Add(this.Rules[i].Clone());
            }

            cf.Pivot = this.Pivot;

            cf.SequenceOfReferences = new List<SLCellPointRange>();
            SLCellPointRange cpr;
            for (i = 0; i < this.SequenceOfReferences.Count; ++i)
            {
                cpr = new SLCellPointRange(this.SequenceOfReferences[i].StartRowIndex, this.SequenceOfReferences[i].StartColumnIndex, this.SequenceOfReferences[i].EndRowIndex, this.SequenceOfReferences[i].EndColumnIndex);
                cf.SequenceOfReferences.Add(cpr);
            }

            return cf;
        }
    }
}
