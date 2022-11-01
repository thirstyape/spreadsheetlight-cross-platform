using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLight
{
    internal class SLIconSet
    {
        internal bool Is2010;

        internal List<SLConditionalFormatValueObject> Cfvos { get; set; }
        internal List<SLConditionalFormattingIcon2010> CustomIcons { get; set; }
        internal SLIconSetValues IconSetType { get; set; }
        internal bool ShowValue { get; set; }
        internal bool Percent { get; set; }
        internal bool Reverse { get; set; }

        internal SLIconSet()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Is2010 = false;
            this.Cfvos = new List<SLConditionalFormatValueObject>();
            this.CustomIcons = new List<SLConditionalFormattingIcon2010>();
            this.IconSetType = SLIconSetValues.ThreeTrafficLights1;
            this.ShowValue = true;
            this.Percent = true;
            this.Reverse = false;
        }

        internal void FromIconSet(IconSet ics)
        {
            this.SetAllNull();

            if (ics.IconSetValue != null) this.IconSetType = SLIconSet.TranslateIconSetToInternalSet(ics.IconSetValue.Value);
            if (ics.ShowValue != null) this.ShowValue = ics.ShowValue.Value;
            if (ics.Percent != null) this.Percent = ics.Percent.Value;
            if (ics.Reverse != null) this.Reverse = ics.Reverse.Value;

            using (OpenXmlReader oxr = OpenXmlReader.Create(ics))
            {
                SLConditionalFormatValueObject cfvo;
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(ConditionalFormatValueObject))
                    {
                        cfvo = new SLConditionalFormatValueObject();
                        cfvo.FromConditionalFormatValueObject((ConditionalFormatValueObject)oxr.LoadCurrentElement());
                        this.Cfvos.Add(cfvo);
                    }
                }
            }
        }

        internal IconSet ToIconSet()
        {
            IconSet ics = new IconSet();
            if (this.IconSetType != SLIconSetValues.ThreeTrafficLights1) ics.IconSetValue = SLIconSet.TranslateInternalSetToIconSet(this.IconSetType);
            if (!this.ShowValue) ics.ShowValue = this.ShowValue;
            if (!this.Percent) ics.Percent = this.Percent;
            if (this.Reverse) ics.Reverse = this.Reverse;

            foreach (SLConditionalFormatValueObject cfvo in this.Cfvos)
            {
                ics.Append(cfvo.ToConditionalFormatValueObject());
            }

            return ics;
        }

        internal SLIconSet2010 ToSLIconSet2010()
        {
            SLIconSet2010 ics2010 = new SLIconSet2010();
            ics2010.IconSetType = SLIconSet.TranslateInternalSetToIconSet2010(this.IconSetType);
            ics2010.ShowValue = this.ShowValue;
            ics2010.Percent = this.Percent;
            ics2010.Reverse = this.Reverse;

            foreach (SLConditionalFormatValueObject cfvo in this.Cfvos)
            {
                ics2010.Cfvos.Add(cfvo.ToSLConditionalFormattingValueObject2010());
            }

            foreach (SLConditionalFormattingIcon2010 cfi in this.CustomIcons)
            {
                ics2010.CustomIcons.Add(cfi.Clone());
            }

            return ics2010;
        }

        internal static SLIconSetValues TranslateIconSetToInternalSet(IconSetValues isv)
        {
            SLIconSetValues result = SLIconSetValues.ThreeTrafficLights1;
            switch (isv)
            {
                case IconSetValues.FiveArrows:
                    result = SLIconSetValues.FiveArrows;
                    break;
                case IconSetValues.FiveArrowsGray:
                    result = SLIconSetValues.FiveArrowsGray;
                    break;
                case IconSetValues.FiveQuarters:
                    result = SLIconSetValues.FiveQuarters;
                    break;
                case IconSetValues.FiveRating:
                    result = SLIconSetValues.FiveRating;
                    break;
                case IconSetValues.FourArrows:
                    result = SLIconSetValues.FourArrows;
                    break;
                case IconSetValues.FourArrowsGray:
                    result = SLIconSetValues.FourArrowsGray;
                    break;
                case IconSetValues.FourRating:
                    result = SLIconSetValues.FourRating;
                    break;
                case IconSetValues.FourRedToBlack:
                    result = SLIconSetValues.FourRedToBlack;
                    break;
                case IconSetValues.FourTrafficLights:
                    result = SLIconSetValues.FourTrafficLights;
                    break;
                case IconSetValues.ThreeArrows:
                    result = SLIconSetValues.ThreeArrows;
                    break;
                case IconSetValues.ThreeArrowsGray:
                    result = SLIconSetValues.ThreeArrowsGray;
                    break;
                case IconSetValues.ThreeFlags:
                    result = SLIconSetValues.ThreeFlags;
                    break;
                case IconSetValues.ThreeSigns:
                    result = SLIconSetValues.ThreeSigns;
                    break;
                case IconSetValues.ThreeSymbols:
                    result = SLIconSetValues.ThreeSymbols;
                    break;
                case IconSetValues.ThreeSymbols2:
                    result = SLIconSetValues.ThreeSymbols2;
                    break;
                case IconSetValues.ThreeTrafficLights1:
                    result = SLIconSetValues.ThreeTrafficLights1;
                    break;
                case IconSetValues.ThreeTrafficLights2:
                    result = SLIconSetValues.ThreeTrafficLights2;
                    break;
            }

            return result;
        }

        internal static IconSetValues TranslateInternalSetToIconSet(SLIconSetValues isv)
        {
            IconSetValues result = IconSetValues.ThreeTrafficLights1;
            switch (isv)
            {
                case SLIconSetValues.FiveArrows:
                    result = IconSetValues.FiveArrows;
                    break;
                case SLIconSetValues.FiveArrowsGray:
                    result = IconSetValues.FiveArrowsGray;
                    break;
                case SLIconSetValues.FiveQuarters:
                    result = IconSetValues.FiveQuarters;
                    break;
                case SLIconSetValues.FiveRating:
                    result = IconSetValues.FiveRating;
                    break;
                case SLIconSetValues.FourArrows:
                    result = IconSetValues.FourArrows;
                    break;
                case SLIconSetValues.FourArrowsGray:
                    result = IconSetValues.FourArrowsGray;
                    break;
                case SLIconSetValues.FourRating:
                    result = IconSetValues.FourRating;
                    break;
                case SLIconSetValues.FourRedToBlack:
                    result = IconSetValues.FourRedToBlack;
                    break;
                case SLIconSetValues.FourTrafficLights:
                    result = IconSetValues.FourTrafficLights;
                    break;
                case SLIconSetValues.ThreeArrows:
                    result = IconSetValues.ThreeArrows;
                    break;
                case SLIconSetValues.ThreeArrowsGray:
                    result = IconSetValues.ThreeArrowsGray;
                    break;
                case SLIconSetValues.ThreeFlags:
                    result = IconSetValues.ThreeFlags;
                    break;
                case SLIconSetValues.ThreeSigns:
                    result = IconSetValues.ThreeSigns;
                    break;
                case SLIconSetValues.ThreeSymbols:
                    result = IconSetValues.ThreeSymbols;
                    break;
                case SLIconSetValues.ThreeSymbols2:
                    result = IconSetValues.ThreeSymbols2;
                    break;
                case SLIconSetValues.ThreeTrafficLights1:
                    result = IconSetValues.ThreeTrafficLights1;
                    break;
                case SLIconSetValues.ThreeTrafficLights2:
                    result = IconSetValues.ThreeTrafficLights2;
                    break;
            }

            return result;
        }

        internal static X14.IconSetTypeValues TranslateInternalSetToIconSet2010(SLIconSetValues isv)
        {
            X14.IconSetTypeValues result = X14.IconSetTypeValues.ThreeTrafficLights1;
            switch (isv)
            {
                case SLIconSetValues.FiveArrows:
                    result = X14.IconSetTypeValues.FiveArrows;
                    break;
                case SLIconSetValues.FiveArrowsGray:
                    result = X14.IconSetTypeValues.FiveArrowsGray;
                    break;
                case SLIconSetValues.FiveBoxes:
                    result = X14.IconSetTypeValues.FiveBoxes;
                    break;
                case SLIconSetValues.FiveQuarters:
                    result = X14.IconSetTypeValues.FiveQuarters;
                    break;
                case SLIconSetValues.FiveRating:
                    result = X14.IconSetTypeValues.FiveRating;
                    break;
                case SLIconSetValues.FourArrows:
                    result = X14.IconSetTypeValues.FourArrows;
                    break;
                case SLIconSetValues.FourArrowsGray:
                    result = X14.IconSetTypeValues.FourArrowsGray;
                    break;
                case SLIconSetValues.FourRating:
                    result = X14.IconSetTypeValues.FourRating;
                    break;
                case SLIconSetValues.FourRedToBlack:
                    result = X14.IconSetTypeValues.FourRedToBlack;
                    break;
                case SLIconSetValues.FourTrafficLights:
                    result = X14.IconSetTypeValues.FourTrafficLights;
                    break;
                case SLIconSetValues.ThreeArrows:
                    result = X14.IconSetTypeValues.ThreeArrows;
                    break;
                case SLIconSetValues.ThreeArrowsGray:
                    result = X14.IconSetTypeValues.ThreeArrowsGray;
                    break;
                case SLIconSetValues.ThreeFlags:
                    result = X14.IconSetTypeValues.ThreeFlags;
                    break;
                case SLIconSetValues.ThreeSigns:
                    result = X14.IconSetTypeValues.ThreeSigns;
                    break;
                case SLIconSetValues.ThreeStars:
                    result = X14.IconSetTypeValues.ThreeStars;
                    break;
                case SLIconSetValues.ThreeSymbols:
                    result = X14.IconSetTypeValues.ThreeSymbols;
                    break;
                case SLIconSetValues.ThreeSymbols2:
                    result = X14.IconSetTypeValues.ThreeSymbols2;
                    break;
                case SLIconSetValues.ThreeTrafficLights1:
                    result = X14.IconSetTypeValues.ThreeTrafficLights1;
                    break;
                case SLIconSetValues.ThreeTrafficLights2:
                    result = X14.IconSetTypeValues.ThreeTrafficLights2;
                    break;
                case SLIconSetValues.ThreeTriangles:
                    result = X14.IconSetTypeValues.ThreeTriangles;
                    break;
            }

            return result;
        }

        internal static bool Is2010IconSet(SLIconSetValues isv)
        {
            bool result = false;

            if (isv == SLIconSetValues.FiveBoxes
                || isv == SLIconSetValues.ThreeStars
                || isv == SLIconSetValues.ThreeTriangles)
            {
                result = true;
            }

            return result;
        }

        internal static void TranslateCustomIcon(SLIconValues Icon, out X14.IconSetTypeValues IconSetType, out uint IconId)
        {
            IconSetType = X14.IconSetTypeValues.ThreeTrafficLights1;
            IconId = 0;

            switch (Icon)
            {
                case SLIconValues.NoIcon:
                    IconSetType = X14.IconSetTypeValues.NoIcons;
                    IconId = 0;
                    break;
                case SLIconValues.GreenUpArrow:
                    IconSetType = X14.IconSetTypeValues.ThreeArrows;
                    IconId = 2;
                    break;
                case SLIconValues.YellowSideArrow:
                    IconSetType = X14.IconSetTypeValues.ThreeArrows;
                    IconId = 1;
                    break;
                case SLIconValues.RedDownArrow:
                    IconSetType = X14.IconSetTypeValues.ThreeArrows;
                    IconId = 0;
                    break;
                case SLIconValues.GrayUpArrow:
                    IconSetType = X14.IconSetTypeValues.ThreeArrowsGray;
                    IconId = 2;
                    break;
                case SLIconValues.GraySideArrow:
                    IconSetType = X14.IconSetTypeValues.ThreeArrowsGray;
                    IconId = 1;
                    break;
                case SLIconValues.GrayDownArrow:
                    IconSetType = X14.IconSetTypeValues.ThreeArrowsGray;
                    IconId = 0;
                    break;
                case SLIconValues.GreenFlag:
                    IconSetType = X14.IconSetTypeValues.ThreeFlags;
                    IconId = 2;
                    break;
                case SLIconValues.YellowFlag:
                    IconSetType = X14.IconSetTypeValues.ThreeFlags;
                    IconId = 1;
                    break;
                case SLIconValues.RedFlag:
                    IconSetType = X14.IconSetTypeValues.ThreeFlags;
                    IconId = 0;
                    break;
                case SLIconValues.GreenCircle:
                    IconSetType = X14.IconSetTypeValues.ThreeTrafficLights1;
                    IconId = 2;
                    break;
                case SLIconValues.YellowCircle:
                    IconSetType = X14.IconSetTypeValues.ThreeTrafficLights1;
                    IconId = 1;
                    break;
                case SLIconValues.RedCircleWithBorder:
                    IconSetType = X14.IconSetTypeValues.ThreeTrafficLights1;
                    IconId = 0;
                    break;
                case SLIconValues.BlackCircleWithBorder:
                    IconSetType = X14.IconSetTypeValues.FourTrafficLights;
                    IconId = 0;
                    break;
                case SLIconValues.GreenTrafficLight:
                    IconSetType = X14.IconSetTypeValues.ThreeTrafficLights2;
                    IconId = 2;
                    break;
                case SLIconValues.YellowTrafficLight:
                    IconSetType = X14.IconSetTypeValues.ThreeTrafficLights2;
                    IconId = 1;
                    break;
                case SLIconValues.RedTrafficLight:
                    IconSetType = X14.IconSetTypeValues.ThreeTrafficLights2;
                    IconId = 0;
                    break;
                case SLIconValues.YellowTriangle:
                    IconSetType = X14.IconSetTypeValues.ThreeSigns;
                    IconId = 1;
                    break;
                case SLIconValues.RedDiamond:
                    IconSetType = X14.IconSetTypeValues.ThreeSigns;
                    IconId = 0;
                    break;
                case SLIconValues.GreenCheckSymbol:
                    IconSetType = X14.IconSetTypeValues.ThreeSymbols;
                    IconId = 2;
                    break;
                case SLIconValues.YellowExclamationSymbol:
                    IconSetType = X14.IconSetTypeValues.ThreeSymbols;
                    IconId = 1;
                    break;
                case SLIconValues.RedCrossSymbol:
                    IconSetType = X14.IconSetTypeValues.ThreeSymbols;
                    IconId = 0;
                    break;
                case SLIconValues.GreenCheck:
                    IconSetType = X14.IconSetTypeValues.ThreeSymbols2;
                    IconId = 2;
                    break;
                case SLIconValues.YellowExclamation:
                    IconSetType = X14.IconSetTypeValues.ThreeSymbols2;
                    IconId = 1;
                    break;
                case SLIconValues.RedCross:
                    IconSetType = X14.IconSetTypeValues.ThreeSymbols2;
                    IconId = 0;
                    break;
                case SLIconValues.YellowUpInclineArrow:
                    IconSetType = X14.IconSetTypeValues.FourArrows;
                    IconId = 2;
                    break;
                case SLIconValues.YellowDownInclineArrow:
                    IconSetType = X14.IconSetTypeValues.FourArrows;
                    IconId = 1;
                    break;
                case SLIconValues.GrayUpInclineArrow:
                    IconSetType = X14.IconSetTypeValues.FourArrowsGray;
                    IconId = 2;
                    break;
                case SLIconValues.GrayDownInclineArrow:
                    IconSetType = X14.IconSetTypeValues.FourArrowsGray;
                    IconId = 1;
                    break;
                case SLIconValues.RedCircle:
                    IconSetType = X14.IconSetTypeValues.FourRedToBlack;
                    IconId = 3;
                    break;
                case SLIconValues.PinkCircle:
                    IconSetType = X14.IconSetTypeValues.FourRedToBlack;
                    IconId = 2;
                    break;
                case SLIconValues.GrayCircle:
                    IconSetType = X14.IconSetTypeValues.FourRedToBlack;
                    IconId = 1;
                    break;
                case SLIconValues.BlackCircle:
                    IconSetType = X14.IconSetTypeValues.FourRedToBlack;
                    IconId = 0;
                    break;
                case SLIconValues.CircleWithOneWhiteQuarter:
                    IconSetType = X14.IconSetTypeValues.FiveQuarters;
                    IconId = 3;
                    break;
                case SLIconValues.CircleWithTwoWhiteQuarters:
                    IconSetType = X14.IconSetTypeValues.FiveQuarters;
                    IconId = 2;
                    break;
                case SLIconValues.CircleWithThreeWhiteQuarters:
                    IconSetType = X14.IconSetTypeValues.FiveQuarters;
                    IconId = 1;
                    break;
                case SLIconValues.WhiteCircleAllWhiteQuarters:
                    IconSetType = X14.IconSetTypeValues.FiveQuarters;
                    IconId = 0;
                    break;
                case SLIconValues.SignalMeterWithNoFilledBars:
                    IconSetType = X14.IconSetTypeValues.FiveRating;
                    IconId = 0;
                    break;
                case SLIconValues.SignalMeterWithOneFilledBar:
                    IconSetType = X14.IconSetTypeValues.FiveRating;
                    IconId = 1;
                    break;
                case SLIconValues.SignalMeterWithTwoFilledBars:
                    IconSetType = X14.IconSetTypeValues.FiveRating;
                    IconId = 2;
                    break;
                case SLIconValues.SignalMeterWithThreeFilledBars:
                    IconSetType = X14.IconSetTypeValues.FiveRating;
                    IconId = 3;
                    break;
                case SLIconValues.SignalMeterWithFourFilledBars:
                    IconSetType = X14.IconSetTypeValues.FiveRating;
                    IconId = 4;
                    break;
                case SLIconValues.GoldStar:
                    IconSetType = X14.IconSetTypeValues.ThreeStars;
                    IconId = 2;
                    break;
                case SLIconValues.HalfGoldStar:
                    IconSetType = X14.IconSetTypeValues.ThreeStars;
                    IconId = 1;
                    break;
                case SLIconValues.SilverStar:
                    IconSetType = X14.IconSetTypeValues.ThreeStars;
                    IconId = 0;
                    break;
                case SLIconValues.GreenUpTriangle:
                    IconSetType = X14.IconSetTypeValues.ThreeTriangles;
                    IconId = 2;
                    break;
                case SLIconValues.YellowDash:
                    IconSetType = X14.IconSetTypeValues.ThreeTriangles;
                    IconId = 1;
                    break;
                case SLIconValues.RedDownTriangle:
                    IconSetType = X14.IconSetTypeValues.ThreeTriangles;
                    IconId = 0;
                    break;
                case SLIconValues.FourFilledBoxes:
                    IconSetType = X14.IconSetTypeValues.FiveBoxes;
                    IconId = 4;
                    break;
                case SLIconValues.ThreeFilledBoxes:
                    IconSetType = X14.IconSetTypeValues.FiveBoxes;
                    IconId = 3;
                    break;
                case SLIconValues.TwoFilledBoxes:
                    IconSetType = X14.IconSetTypeValues.FiveBoxes;
                    IconId = 2;
                    break;
                case SLIconValues.OneFilledBox:
                    IconSetType = X14.IconSetTypeValues.FiveBoxes;
                    IconId = 1;
                    break;
                case SLIconValues.ZeroFilledBoxes:
                    IconSetType = X14.IconSetTypeValues.FiveBoxes;
                    IconId = 0;
                    break;
            }
        }

        internal static SLIconValues TranslateIconSetType(X14.IconSetTypeValues IconSetType, uint IconId)
        {
            SLIconValues result = SLIconValues.NoIcon;
            switch (IconSetType)
            {
                case X14.IconSetTypeValues.FiveArrows:
                    if (IconId == 0) result = SLIconValues.RedDownArrow;
                    else if (IconId == 1) result = SLIconValues.YellowDownInclineArrow;
                    else if (IconId == 2) result = SLIconValues.YellowSideArrow;
                    else if (IconId == 3) result = SLIconValues.YellowUpInclineArrow;
                    else if (IconId == 4) result = SLIconValues.GreenUpArrow;
                    break;
                case X14.IconSetTypeValues.FiveArrowsGray:
                    if (IconId == 0) result = SLIconValues.GrayDownArrow;
                    else if (IconId == 1) result = SLIconValues.GrayDownInclineArrow;
                    else if (IconId == 2) result = SLIconValues.GraySideArrow;
                    else if (IconId == 3) result = SLIconValues.GrayUpInclineArrow;
                    else if (IconId == 4) result = SLIconValues.GrayUpArrow;
                    break;
                case X14.IconSetTypeValues.FiveBoxes:
                    if (IconId == 0) result = SLIconValues.ZeroFilledBoxes;
                    else if (IconId == 1) result = SLIconValues.OneFilledBox;
                    else if (IconId == 2) result = SLIconValues.TwoFilledBoxes;
                    else if (IconId == 3) result = SLIconValues.ThreeFilledBoxes;
                    else if (IconId == 4) result = SLIconValues.FourFilledBoxes;
                    break;
                case X14.IconSetTypeValues.FiveQuarters:
                    if (IconId == 0) result = SLIconValues.WhiteCircleAllWhiteQuarters;
                    else if (IconId == 1) result = SLIconValues.CircleWithThreeWhiteQuarters;
                    else if (IconId == 2) result = SLIconValues.CircleWithTwoWhiteQuarters;
                    else if (IconId == 3) result = SLIconValues.CircleWithOneWhiteQuarter;
                    else if (IconId == 4) result = SLIconValues.BlackCircle;
                    break;
                case X14.IconSetTypeValues.FiveRating:
                    if (IconId == 0) result = SLIconValues.SignalMeterWithNoFilledBars;
                    else if (IconId == 1) result = SLIconValues.SignalMeterWithOneFilledBar;
                    else if (IconId == 2) result = SLIconValues.SignalMeterWithTwoFilledBars;
                    else if (IconId == 3) result = SLIconValues.SignalMeterWithThreeFilledBars;
                    else if (IconId == 4) result = SLIconValues.SignalMeterWithFourFilledBars;
                    break;
                case X14.IconSetTypeValues.FourArrows:
                    if (IconId == 0) result = SLIconValues.RedDownArrow;
                    else if (IconId == 1) result = SLIconValues.YellowDownInclineArrow;
                    else if (IconId == 2) result = SLIconValues.YellowUpInclineArrow;
                    else if (IconId == 3) result = SLIconValues.GreenUpArrow;
                    break;
                case X14.IconSetTypeValues.FourArrowsGray:
                    if (IconId == 0) result = SLIconValues.GrayDownArrow;
                    else if (IconId == 1) result = SLIconValues.GrayDownInclineArrow;
                    else if (IconId == 2) result = SLIconValues.GrayUpInclineArrow;
                    else if (IconId == 3) result = SLIconValues.GrayUpArrow;
                    break;
                case X14.IconSetTypeValues.FourRating:
                    if (IconId == 0) result = SLIconValues.SignalMeterWithOneFilledBar;
                    else if (IconId == 1) result = SLIconValues.SignalMeterWithTwoFilledBars;
                    else if (IconId == 2) result = SLIconValues.SignalMeterWithThreeFilledBars;
                    else if (IconId == 3) result = SLIconValues.SignalMeterWithFourFilledBars;
                    break;
                case X14.IconSetTypeValues.FourRedToBlack:
                    if (IconId == 0) result = SLIconValues.BlackCircle;
                    else if (IconId == 1) result = SLIconValues.GrayCircle;
                    else if (IconId == 2) result = SLIconValues.PinkCircle;
                    else if (IconId == 3) result = SLIconValues.RedCircle;
                    break;
                case X14.IconSetTypeValues.FourTrafficLights:
                    if (IconId == 0) result = SLIconValues.BlackCircleWithBorder;
                    else if (IconId == 1) result = SLIconValues.RedCircleWithBorder;
                    else if (IconId == 2) result = SLIconValues.YellowCircle;
                    else if (IconId == 3) result = SLIconValues.GreenCircle;
                    break;
                case X14.IconSetTypeValues.NoIcons:
                    result = SLIconValues.NoIcon;
                    break;
                case X14.IconSetTypeValues.ThreeArrows:
                    if (IconId == 0) result = SLIconValues.RedDownArrow;
                    else if (IconId == 1) result = SLIconValues.YellowSideArrow;
                    else if (IconId == 2) result = SLIconValues.GreenUpArrow;
                    break;
                case X14.IconSetTypeValues.ThreeArrowsGray:
                    if (IconId == 0) result = SLIconValues.GrayDownArrow;
                    else if (IconId == 1) result = SLIconValues.GraySideArrow;
                    else if (IconId == 2) result = SLIconValues.GrayUpArrow;
                    break;
                case X14.IconSetTypeValues.ThreeFlags:
                    if (IconId == 0) result = SLIconValues.RedFlag;
                    else if (IconId == 1) result = SLIconValues.YellowFlag;
                    else if (IconId == 2) result = SLIconValues.GreenFlag;
                    break;
                case X14.IconSetTypeValues.ThreeSigns:
                    if (IconId == 0) result = SLIconValues.RedDiamond;
                    else if (IconId == 1) result = SLIconValues.YellowTriangle;
                    else if (IconId == 2) result = SLIconValues.GreenCircle;
                    break;
                case X14.IconSetTypeValues.ThreeStars:
                    if (IconId == 0) result = SLIconValues.SilverStar;
                    else if (IconId == 1) result = SLIconValues.HalfGoldStar;
                    else if (IconId == 2) result = SLIconValues.GoldStar;
                    break;
                case X14.IconSetTypeValues.ThreeSymbols:
                    if (IconId == 0) result = SLIconValues.RedCrossSymbol;
                    else if (IconId == 1) result = SLIconValues.YellowExclamationSymbol;
                    else if (IconId == 2) result = SLIconValues.GreenCheckSymbol;
                    break;
                case X14.IconSetTypeValues.ThreeSymbols2:
                    if (IconId == 0) result = SLIconValues.RedCross;
                    else if (IconId == 1) result = SLIconValues.YellowExclamation;
                    else if (IconId == 2) result = SLIconValues.GreenCheck;
                    break;
                case X14.IconSetTypeValues.ThreeTrafficLights1:
                    if (IconId == 0) result = SLIconValues.RedCircleWithBorder;
                    else if (IconId == 1) result = SLIconValues.YellowCircle;
                    else if (IconId == 2) result = SLIconValues.GreenCircle;
                    break;
                case X14.IconSetTypeValues.ThreeTrafficLights2:
                    if (IconId == 0) result = SLIconValues.RedTrafficLight;
                    else if (IconId == 1) result = SLIconValues.YellowTrafficLight;
                    else if (IconId == 2) result = SLIconValues.GreenTrafficLight;
                    break;
                case X14.IconSetTypeValues.ThreeTriangles:
                    if (IconId == 0) result = SLIconValues.RedDownArrow;
                    else if (IconId == 1) result = SLIconValues.YellowDash;
                    else if (IconId == 2) result = SLIconValues.GreenUpArrow;
                    break;
            }

            return result;
        }

        internal SLIconSet Clone()
        {
            SLIconSet ics = new SLIconSet();
            ics.Cfvos = new List<SLConditionalFormatValueObject>();
            for (int i = 0; i < this.Cfvos.Count; ++i)
            {
                ics.Cfvos.Add(this.Cfvos[i].Clone());
            }

            ics.IconSetType = this.IconSetType;
            ics.ShowValue = this.ShowValue;
            ics.Percent = this.Percent;
            ics.Reverse = this.Reverse;

            return ics;
        }
    }
}
