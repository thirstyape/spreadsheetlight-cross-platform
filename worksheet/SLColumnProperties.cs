using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for columns. This simulates the DocumentFormat.OpenXml.Spreadsheet.Column class.
    /// </summary>
    internal class SLColumnProperties
    {
        internal bool IsEmpty
        {
            get
            {
                return !this.HasWidth && (this.StyleIndex == 0) && !this.Hidden
                    && !this.BestFit && !this.Phonetic && (this.OutlineLevel == 0)
                    && !this.Collapsed;
            }
        }

        internal int MaxDigitWidth { get; set; }
        internal List<double> listColumnStepSize { get; set; }
        internal double ThemeDefaultColumnWidth;
        internal long ThemeDefaultColumnWidthInEMU;

        // this doubles as customWidth
        internal bool HasWidth;
        private double fWidth;
        // The column width. This is in number of characters of the width of the digit (0, 1, ... 9) with the maximum width, as rendered in the Normal style's font. The Normal style's font is typically the minor font in the default point size.
        internal double Width
        {
            get { return fWidth; }
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
                    lWidthInEMU = (long)iPixels * SLDocument.PixelToEMU;
                    fWidth = iWholeNumber + this.listColumnStepSize[iStep];
                    HasWidth = true;

                    this.BestFit = false;
                }
            }
        }

        private long lWidthInEMU;
        internal long WidthInEMU
        {
            get { return lWidthInEMU; }
        }

        internal uint StyleIndex { get; set; }
        internal bool Hidden { get; set; }
        internal bool BestFit { get; set; }
        internal bool Phonetic { get; set; }
        internal byte OutlineLevel { get; set; }
        internal bool Collapsed { get; set; }

        /// <summary>
        /// Initializes an instance of SLColumnProperties.
        /// </summary>
        internal SLColumnProperties(double ThemeDefaultColumnWidth, long ThemeDefaultColumnWidthInEMU, int MaxDigitWidth, List<double> ColumnStepSize)
        {
            this.MaxDigitWidth = MaxDigitWidth;
            this.listColumnStepSize = new List<double>();
            for (int i = 0; i < ColumnStepSize.Count; ++i)
            {
                this.listColumnStepSize.Add(ColumnStepSize[i]);
            }

            this.ThemeDefaultColumnWidth = ThemeDefaultColumnWidth;
            this.ThemeDefaultColumnWidthInEMU = ThemeDefaultColumnWidthInEMU;
            this.Width = ThemeDefaultColumnWidth;
            this.lWidthInEMU = ThemeDefaultColumnWidthInEMU;
            this.HasWidth = false;

            this.StyleIndex = 0;
            this.Hidden = false;
            this.BestFit = false;
            this.Phonetic = false;
            this.OutlineLevel = 0;
            this.Collapsed = false;
        }

        internal string ToHash()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("{0},", this.HasWidth);
            sb.AppendFormat("{0},", this.Width.ToString(CultureInfo.InvariantCulture));
            sb.AppendFormat("{0},", this.StyleIndex.ToString(CultureInfo.InvariantCulture));
            sb.AppendFormat("{0},", this.Hidden);
            sb.AppendFormat("{0},", this.BestFit);
            sb.AppendFormat("{0},", this.Phonetic);
            sb.AppendFormat("{0},", this.OutlineLevel.ToString(CultureInfo.InvariantCulture));
            sb.AppendFormat("{0}", this.Collapsed);

            return sb.ToString();
        }

        internal SLColumnProperties Clone()
        {
            SLColumnProperties cp = new SLColumnProperties(this.ThemeDefaultColumnWidth, this.ThemeDefaultColumnWidthInEMU, this.MaxDigitWidth, this.listColumnStepSize);
            cp.HasWidth = this.HasWidth;
            cp.fWidth = this.fWidth;
            cp.lWidthInEMU = this.lWidthInEMU;
            cp.StyleIndex = this.StyleIndex;
            cp.Hidden = this.Hidden;
            cp.BestFit = this.BestFit;
            cp.Phonetic = this.Phonetic;
            cp.OutlineLevel = this.OutlineLevel;
            cp.Collapsed = this.Collapsed;

            return cp;
        }
    }
}
