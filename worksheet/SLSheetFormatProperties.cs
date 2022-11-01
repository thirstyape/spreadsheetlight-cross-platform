using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLSheetFormatProperties
    {
        internal int MaxDigitWidth { get; set; }
        internal List<double> listColumnStepSize { get; set; }
        internal double ThemeDefaultColumnWidth;
        internal long ThemeDefaultColumnWidthInEMU;

        internal uint? BaseColumnWidth { get; set; }

        internal bool HasDefaultColumnWidth { get; set; }

        internal double fDefaultColumnWidth;
        internal double DefaultColumnWidth
        {
            get
            {
                return fDefaultColumnWidth;
            }
            set
            {
                double fValue = value;
                if (fValue > 0)
                {
                    int iWholeNumber = Convert.ToInt32(Math.Truncate(fValue));
                    double fRemainder = fValue - (double)iWholeNumber;

                    int iStep = 0;
                    for (iStep = listColumnStepSize.Count - 1; iStep >= 0; --iStep)
                    {
                        if (fRemainder > listColumnStepSize[iStep]) break;
                    }

                    // this is in case (fRemainder > listColumnStepSize[iStep]) evaluates
                    // to false when fRemainder is 0.0 and listColumnStepSize[0] is also 0.0
                    // and I hate checking for equality between floating point values...
                    // By then iStep should be -1, which breaks the loop.
                    if (iStep < 0) iStep = 0;

                    // the step sizes were calculated based on the max digit width minus 1 pixel.
                    int iPixels = iWholeNumber * (this.MaxDigitWidth - 1) + iStep;
                    lDefaultColumnWidthInEMU = (long)iPixels * SLDocument.PixelToEMU;
                    fDefaultColumnWidth = iWholeNumber + this.listColumnStepSize[iStep];
                    HasDefaultColumnWidth = true;
                }
            }
        }

        internal long lDefaultColumnWidthInEMU;
        internal long DefaultColumnWidthInEMU
        {
            get
            {
                return lDefaultColumnWidthInEMU;
            }
        }

        internal double CalculatedDefaultRowHeight;

        internal double fDefaultRowHeight;
        internal double DefaultRowHeight
        {
            get
            {
                return fDefaultRowHeight;
            }
            set
            {
                double fModifiedRowHeight = value / SLDocument.RowHeightMultiple;
                // round because it looks nicer. Is 4 decimal places good enough?
                fModifiedRowHeight = Math.Round(Math.Ceiling(fModifiedRowHeight) * SLDocument.RowHeightMultiple, 4);

                lDefaultRowHeightInEMU = (long)(fModifiedRowHeight * SLConstants.PointToEMU);

                fDefaultRowHeight = fModifiedRowHeight;
            }
        }

        internal long lDefaultRowHeightInEMU;
        internal long DefaultRowHeightInEMU
        {
            get
            {
                return lDefaultRowHeightInEMU;
            }
        }

        internal bool? CustomHeight { get; set; }
        internal bool? ZeroHeight { get; set; }
        internal bool? ThickTop { get; set; }
        internal bool? ThickBottom { get; set; }
        internal byte? OutlineLevelRow { get; set; }
        internal byte? OutlineLevelColumn { get; set; }

        internal SLSheetFormatProperties(double ThemeDefaultColumnWidth, long ThemeDefaultColumnWidthInEMU, int MaxDigitWidth, List<double> ColumnStepSize, double CalculatedDefaultRowHeight)
        {
            this.MaxDigitWidth = MaxDigitWidth;
            this.listColumnStepSize = new List<double>();
            for (int i = 0; i < ColumnStepSize.Count; ++i)
            {
                this.listColumnStepSize.Add(ColumnStepSize[i]);
            }

            this.BaseColumnWidth = null;

            this.ThemeDefaultColumnWidth = ThemeDefaultColumnWidth;
            this.ThemeDefaultColumnWidthInEMU = ThemeDefaultColumnWidthInEMU;
            this.fDefaultColumnWidth = ThemeDefaultColumnWidth;
            this.lDefaultColumnWidthInEMU = ThemeDefaultColumnWidthInEMU;
            this.HasDefaultColumnWidth = false;

            this.CalculatedDefaultRowHeight = CalculatedDefaultRowHeight;
            this.fDefaultRowHeight = CalculatedDefaultRowHeight;
            this.lDefaultRowHeightInEMU = Convert.ToInt64(CalculatedDefaultRowHeight * SLConstants.PointToEMU);

            this.CustomHeight = null;
            this.ZeroHeight = null;
            this.ThickTop = null;
            this.ThickBottom = null;
            this.OutlineLevelRow = null;
            this.OutlineLevelColumn = null;
        }

        internal void FromSheetFormatProperties(SheetFormatProperties sfp)
        {
            if (sfp.BaseColumnWidth != null) this.BaseColumnWidth = sfp.BaseColumnWidth.Value;
            else this.BaseColumnWidth = null;

            if (sfp.DefaultColumnWidth != null)
            {
                this.DefaultColumnWidth = sfp.DefaultColumnWidth.Value;
                this.HasDefaultColumnWidth = true;
            }
            else
            {
                this.fDefaultColumnWidth = this.ThemeDefaultColumnWidth;
                this.lDefaultRowHeightInEMU = this.ThemeDefaultColumnWidthInEMU;
                this.HasDefaultColumnWidth = false;
            }

            if (sfp.DefaultRowHeight != null)
            {
                this.DefaultRowHeight = sfp.DefaultRowHeight.Value;
            }
            else
            {
                this.fDefaultRowHeight = this.CalculatedDefaultRowHeight;
                this.lDefaultRowHeightInEMU = Convert.ToInt64(this.CalculatedDefaultRowHeight * SLConstants.PointToEMU);
            }

            if (sfp.CustomHeight != null) this.CustomHeight = sfp.CustomHeight.Value;
            else this.CustomHeight = null;

            if (sfp.ZeroHeight != null) this.ZeroHeight = sfp.ZeroHeight.Value;
            else this.ZeroHeight = null;

            if (sfp.ThickTop != null) this.ThickTop = sfp.ThickTop.Value;
            else this.ThickTop = null;

            if (sfp.ThickBottom != null) this.ThickBottom = sfp.ThickBottom.Value;
            else this.ThickBottom = null;

            if (sfp.OutlineLevelRow != null) this.OutlineLevelRow = sfp.OutlineLevelRow.Value;
            else this.OutlineLevelRow = null;

            if (sfp.OutlineLevelColumn != null) this.OutlineLevelColumn = sfp.OutlineLevelColumn.Value;
            else this.OutlineLevelColumn = null;
        }

        internal SheetFormatProperties ToSheetFormatProperties()
        {
            SheetFormatProperties sfp = new SheetFormatProperties();
            if (this.BaseColumnWidth != null) sfp.BaseColumnWidth = this.BaseColumnWidth.Value;

            if (this.HasDefaultColumnWidth)
            {
                sfp.DefaultColumnWidth = this.DefaultColumnWidth;
            }

            sfp.DefaultRowHeight = this.DefaultRowHeight;

            if (this.CustomHeight != null) sfp.CustomHeight = this.CustomHeight.Value;
            if (this.ZeroHeight != null) sfp.ZeroHeight = this.ZeroHeight.Value;
            if (this.ThickTop != null) sfp.ThickTop = this.ThickTop.Value;
            if (this.ThickBottom != null) sfp.ThickBottom = this.ThickBottom.Value;
            if (this.OutlineLevelRow != null) sfp.OutlineLevelRow = this.OutlineLevelRow.Value;
            if (this.OutlineLevelColumn != null) sfp.OutlineLevelColumn = this.OutlineLevelColumn.Value;

            return sfp;
        }
    }
}
