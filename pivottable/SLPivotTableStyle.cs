using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    public enum SLPivotTableStyleTypeValues
    {
        /// <summary>
        /// Pivot Style Light 1
        /// </summary>
        Light1 = 0,
        /// <summary>
        /// Pivot Style Light 2
        /// </summary>
        Light2,
        /// <summary>
        /// Pivot Style Light 3
        /// </summary>
        Light3,
        /// <summary>
        /// Pivot Style Light 4
        /// </summary>
        Light4,
        /// <summary>
        /// Pivot Style Light 5
        /// </summary>
        Light5,
        /// <summary>
        /// Pivot Style Light 6
        /// </summary>
        Light6,
        /// <summary>
        /// Pivot Style Light 7
        /// </summary>
        Light7,
        /// <summary>
        /// Pivot Style Light 8
        /// </summary>
        Light8,
        /// <summary>
        /// Pivot Style Light 9
        /// </summary>
        Light9,
        /// <summary>
        /// Pivot Style Light 10
        /// </summary>
        Light10,
        /// <summary>
        /// Pivot Style Light 11
        /// </summary>
        Light11,
        /// <summary>
        /// Pivot Style Light 12
        /// </summary>
        Light12,
        /// <summary>
        /// Pivot Style Light 13
        /// </summary>
        Light13,
        /// <summary>
        /// Pivot Style Light 14
        /// </summary>
        Light14,
        /// <summary>
        /// Pivot Style Light 15
        /// </summary>
        Light15,
        /// <summary>
        /// Pivot Style Light 16
        /// </summary>
        Light16,
        /// <summary>
        /// Pivot Style Light 17
        /// </summary>
        Light17,
        /// <summary>
        /// Pivot Style Light 18
        /// </summary>
        Light18,
        /// <summary>
        /// Pivot Style Light 19
        /// </summary>
        Light19,
        /// <summary>
        /// Pivot Style Light 20
        /// </summary>
        Light20,
        /// <summary>
        /// Pivot Style Light 21
        /// </summary>
        Light21,
        /// <summary>
        /// Pivot Style Light 22
        /// </summary>
        Light22,
        /// <summary>
        /// Pivot Style Light 23
        /// </summary>
        Light23,
        /// <summary>
        /// Pivot Style Light 24
        /// </summary>
        Light24,
        /// <summary>
        /// Pivot Style Light 25
        /// </summary>
        Light25,
        /// <summary>
        /// Pivot Style Light 26
        /// </summary>
        Light26,
        /// <summary>
        /// Pivot Style Light 27
        /// </summary>
        Light27,
        /// <summary>
        /// Pivot Style Light 28
        /// </summary>
        Light28,
        /// <summary>
        /// Pivot Style Medium 1
        /// </summary>
        Medium1,
        /// <summary>
        /// Pivot Style Medium 2
        /// </summary>
        Medium2,
        /// <summary>
        /// Pivot Style Medium 3
        /// </summary>
        Medium3,
        /// <summary>
        /// Pivot Style Medium 4
        /// </summary>
        Medium4,
        /// <summary>
        /// Pivot Style Medium 5
        /// </summary>
        Medium5,
        /// <summary>
        /// Pivot Style Medium 6
        /// </summary>
        Medium6,
        /// <summary>
        /// Pivot Style Medium 7
        /// </summary>
        Medium7,
        /// <summary>
        /// Pivot Style Medium 8
        /// </summary>
        Medium8,
        /// <summary>
        /// Pivot Style Medium 9
        /// </summary>
        Medium9,
        /// <summary>
        /// Pivot Style Medium 10
        /// </summary>
        Medium10,
        /// <summary>
        /// Pivot Style Medium 11
        /// </summary>
        Medium11,
        /// <summary>
        /// Pivot Style Medium 12
        /// </summary>
        Medium12,
        /// <summary>
        /// Pivot Style Medium 13
        /// </summary>
        Medium13,
        /// <summary>
        /// Pivot Style Medium 14
        /// </summary>
        Medium14,
        /// <summary>
        /// Pivot Style Medium 15
        /// </summary>
        Medium15,
        /// <summary>
        /// Pivot Style Medium 16
        /// </summary>
        Medium16,
        /// <summary>
        /// Pivot Style Medium 17
        /// </summary>
        Medium17,
        /// <summary>
        /// Pivot Style Medium 18
        /// </summary>
        Medium18,
        /// <summary>
        /// Pivot Style Medium 19
        /// </summary>
        Medium19,
        /// <summary>
        /// Pivot Style Medium 20
        /// </summary>
        Medium20,
        /// <summary>
        /// Pivot Style Medium 21
        /// </summary>
        Medium21,
        /// <summary>
        /// Pivot Style Medium 22
        /// </summary>
        Medium22,
        /// <summary>
        /// Pivot Style Medium 23
        /// </summary>
        Medium23,
        /// <summary>
        /// Pivot Style Medium 24
        /// </summary>
        Medium24,
        /// <summary>
        /// Pivot Style Medium 25
        /// </summary>
        Medium25,
        /// <summary>
        /// Pivot Style Medium 26
        /// </summary>
        Medium26,
        /// <summary>
        /// Pivot Style Medium 27
        /// </summary>
        Medium27,
        /// <summary>
        /// Pivot Style Medium 28
        /// </summary>
        Medium28,
        /// <summary>
        /// Pivot Style Dark 1
        /// </summary>
        Dark1,
        /// <summary>
        /// Pivot Style Dark 2
        /// </summary>
        Dark2,
        /// <summary>
        /// Pivot Style Dark 3
        /// </summary>
        Dark3,
        /// <summary>
        /// Pivot Style Dark 4
        /// </summary>
        Dark4,
        /// <summary>
        /// Pivot Style Dark 5
        /// </summary>
        Dark5,
        /// <summary>
        /// Pivot Style Dark 6
        /// </summary>
        Dark6,
        /// <summary>
        /// Pivot Style Dark 7
        /// </summary>
        Dark7,
        /// <summary>
        /// Pivot Style Dark 8
        /// </summary>
        Dark8,
        /// <summary>
        /// Pivot Style Dark 9
        /// </summary>
        Dark9,
        /// <summary>
        /// Pivot Style Dark 10
        /// </summary>
        Dark10,
        /// <summary>
        /// Pivot Style Dark 11
        /// </summary>
        Dark11,
        /// <summary>
        /// Pivot Style Dark 12
        /// </summary>
        Dark12,
        /// <summary>
        /// Pivot Style Dark 13
        /// </summary>
        Dark13,
        /// <summary>
        /// Pivot Style Dark 14
        /// </summary>
        Dark14,
        /// <summary>
        /// Pivot Style Dark 15
        /// </summary>
        Dark15,
        /// <summary>
        /// Pivot Style Dark 16
        /// </summary>
        Dark16,
        /// <summary>
        /// Pivot Style Dark 17
        /// </summary>
        Dark17,
        /// <summary>
        /// Pivot Style Dark 18
        /// </summary>
        Dark18,
        /// <summary>
        /// Pivot Style Dark 19
        /// </summary>
        Dark19,
        /// <summary>
        /// Pivot Style Dark 20
        /// </summary>
        Dark20,
        /// <summary>
        /// Pivot Style Dark 21
        /// </summary>
        Dark21,
        /// <summary>
        /// Pivot Style Dark 22
        /// </summary>
        Dark22,
        /// <summary>
        /// Pivot Style Dark 23
        /// </summary>
        Dark23,
        /// <summary>
        /// Pivot Style Dark 24
        /// </summary>
        Dark24,
        /// <summary>
        /// Pivot Style Dark 25
        /// </summary>
        Dark25,
        /// <summary>
        /// Pivot Style Dark 26
        /// </summary>
        Dark26,
        /// <summary>
        /// Pivot Style Dark 27
        /// </summary>
        Dark27,
        /// <summary>
        /// Pivot Style Dark 28
        /// </summary>
        Dark28
    }

    internal class SLPivotTableStyle
    {
        internal string Name { get; set; }
        internal bool ShowRowHeaders { get; set; }
        internal bool ShowColumnHeaders { get; set; }
        internal bool ShowRowStripes { get; set; }
        internal bool ShowColumnStripes { get; set; }
        internal bool? ShowLastColumn { get; set; }

        internal SLPivotTableStyle()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Name = SLConstants.DefaultPivotStyle;
            this.ShowRowHeaders = true;
            this.ShowColumnHeaders = true;
            this.ShowRowStripes = false;
            this.ShowColumnStripes = false;
            this.ShowLastColumn = true;
        }

        internal void FromPivotTableStyle(PivotTableStyle pts)
        {
            this.SetAllNull();

            if (pts.Name != null) this.Name = pts.Name.Value;
            if (pts.ShowRowHeaders != null) this.ShowRowHeaders = pts.ShowRowHeaders.Value;
            if (pts.ShowColumnHeaders != null) this.ShowColumnHeaders = pts.ShowColumnHeaders.Value;
            if (pts.ShowRowStripes != null) this.ShowRowStripes = pts.ShowRowStripes.Value;
            if (pts.ShowColumnStripes != null) this.ShowColumnStripes = pts.ShowColumnStripes.Value;
            if (pts.ShowLastColumn != null) this.ShowLastColumn = pts.ShowLastColumn.Value;
        }

        internal PivotTableStyle ToPivotTableStyle()
        {
            PivotTableStyle pts = new PivotTableStyle();
            if (this.Name != null && this.Name.Length > 0) pts.Name = this.Name;
            pts.ShowRowHeaders = this.ShowRowHeaders;
            pts.ShowColumnHeaders = this.ShowColumnHeaders;
            pts.ShowRowStripes = this.ShowRowStripes;
            pts.ShowColumnStripes = this.ShowColumnStripes;
            if (this.ShowLastColumn != null) pts.ShowLastColumn = this.ShowLastColumn.Value;

            return pts;
        }

        internal void SetPivotTableStyle(SLPivotTableStyleTypeValues pivotstyle)
        {
            switch (pivotstyle)
            {
                case SLPivotTableStyleTypeValues.Light1:
                    this.Name = "PivotStyleLight1";
                    break;
                case SLPivotTableStyleTypeValues.Light2:
                    this.Name = "PivotStyleLight2";
                    break;
                case SLPivotTableStyleTypeValues.Light3:
                    this.Name = "PivotStyleLight3";
                    break;
                case SLPivotTableStyleTypeValues.Light4:
                    this.Name = "PivotStyleLight4";
                    break;
                case SLPivotTableStyleTypeValues.Light5:
                    this.Name = "PivotStyleLight5";
                    break;
                case SLPivotTableStyleTypeValues.Light6:
                    this.Name = "PivotStyleLight6";
                    break;
                case SLPivotTableStyleTypeValues.Light7:
                    this.Name = "PivotStyleLight7";
                    break;
                case SLPivotTableStyleTypeValues.Light8:
                    this.Name = "PivotStyleLight8";
                    break;
                case SLPivotTableStyleTypeValues.Light9:
                    this.Name = "PivotStyleLight9";
                    break;
                case SLPivotTableStyleTypeValues.Light10:
                    this.Name = "PivotStyleLight10";
                    break;
                case SLPivotTableStyleTypeValues.Light11:
                    this.Name = "PivotStyleLight11";
                    break;
                case SLPivotTableStyleTypeValues.Light12:
                    this.Name = "PivotStyleLight12";
                    break;
                case SLPivotTableStyleTypeValues.Light13:
                    this.Name = "PivotStyleLight13";
                    break;
                case SLPivotTableStyleTypeValues.Light14:
                    this.Name = "PivotStyleLight14";
                    break;
                case SLPivotTableStyleTypeValues.Light15:
                    this.Name = "PivotStyleLight15";
                    break;
                case SLPivotTableStyleTypeValues.Light16:
                    this.Name = "PivotStyleLight16";
                    break;
                case SLPivotTableStyleTypeValues.Light17:
                    this.Name = "PivotStyleLight17";
                    break;
                case SLPivotTableStyleTypeValues.Light18:
                    this.Name = "PivotStyleLight18";
                    break;
                case SLPivotTableStyleTypeValues.Light19:
                    this.Name = "PivotStyleLight19";
                    break;
                case SLPivotTableStyleTypeValues.Light20:
                    this.Name = "PivotStyleLight20";
                    break;
                case SLPivotTableStyleTypeValues.Light21:
                    this.Name = "PivotStyleLight21";
                    break;
                case SLPivotTableStyleTypeValues.Light22:
                    this.Name = "PivotStyleLight22";
                    break;
                case SLPivotTableStyleTypeValues.Light23:
                    this.Name = "PivotStyleLight23";
                    break;
                case SLPivotTableStyleTypeValues.Light24:
                    this.Name = "PivotStyleLight24";
                    break;
                case SLPivotTableStyleTypeValues.Light25:
                    this.Name = "PivotStyleLight25";
                    break;
                case SLPivotTableStyleTypeValues.Light26:
                    this.Name = "PivotStyleLight26";
                    break;
                case SLPivotTableStyleTypeValues.Light27:
                    this.Name = "PivotStyleLight27";
                    break;
                case SLPivotTableStyleTypeValues.Light28:
                    this.Name = "PivotStyleLight28";
                    break;
                case SLPivotTableStyleTypeValues.Medium1:
                    this.Name = "PivotStyleMedium1";
                    break;
                case SLPivotTableStyleTypeValues.Medium2:
                    this.Name = "PivotStyleMedium2";
                    break;
                case SLPivotTableStyleTypeValues.Medium3:
                    this.Name = "PivotStyleMedium3";
                    break;
                case SLPivotTableStyleTypeValues.Medium4:
                    this.Name = "PivotStyleMedium4";
                    break;
                case SLPivotTableStyleTypeValues.Medium5:
                    this.Name = "PivotStyleMedium5";
                    break;
                case SLPivotTableStyleTypeValues.Medium6:
                    this.Name = "PivotStyleMedium6";
                    break;
                case SLPivotTableStyleTypeValues.Medium7:
                    this.Name = "PivotStyleMedium7";
                    break;
                case SLPivotTableStyleTypeValues.Medium8:
                    this.Name = "PivotStyleMedium8";
                    break;
                case SLPivotTableStyleTypeValues.Medium9:
                    this.Name = "PivotStyleMedium9";
                    break;
                case SLPivotTableStyleTypeValues.Medium10:
                    this.Name = "PivotStyleMedium10";
                    break;
                case SLPivotTableStyleTypeValues.Medium11:
                    this.Name = "PivotStyleMedium11";
                    break;
                case SLPivotTableStyleTypeValues.Medium12:
                    this.Name = "PivotStyleMedium12";
                    break;
                case SLPivotTableStyleTypeValues.Medium13:
                    this.Name = "PivotStyleMedium13";
                    break;
                case SLPivotTableStyleTypeValues.Medium14:
                    this.Name = "PivotStyleMedium14";
                    break;
                case SLPivotTableStyleTypeValues.Medium15:
                    this.Name = "PivotStyleMedium15";
                    break;
                case SLPivotTableStyleTypeValues.Medium16:
                    this.Name = "PivotStyleMedium16";
                    break;
                case SLPivotTableStyleTypeValues.Medium17:
                    this.Name = "PivotStyleMedium17";
                    break;
                case SLPivotTableStyleTypeValues.Medium18:
                    this.Name = "PivotStyleMedium18";
                    break;
                case SLPivotTableStyleTypeValues.Medium19:
                    this.Name = "PivotStyleMedium19";
                    break;
                case SLPivotTableStyleTypeValues.Medium20:
                    this.Name = "PivotStyleMedium20";
                    break;
                case SLPivotTableStyleTypeValues.Medium21:
                    this.Name = "PivotStyleMedium21";
                    break;
                case SLPivotTableStyleTypeValues.Medium22:
                    this.Name = "PivotStyleMedium22";
                    break;
                case SLPivotTableStyleTypeValues.Medium23:
                    this.Name = "PivotStyleMedium23";
                    break;
                case SLPivotTableStyleTypeValues.Medium24:
                    this.Name = "PivotStyleMedium24";
                    break;
                case SLPivotTableStyleTypeValues.Medium25:
                    this.Name = "PivotStyleMedium25";
                    break;
                case SLPivotTableStyleTypeValues.Medium26:
                    this.Name = "PivotStyleMedium26";
                    break;
                case SLPivotTableStyleTypeValues.Medium27:
                    this.Name = "PivotStyleMedium27";
                    break;
                case SLPivotTableStyleTypeValues.Medium28:
                    this.Name = "PivotStyleMedium28";
                    break;
                case SLPivotTableStyleTypeValues.Dark1:
                    this.Name = "PivotStyleDark1";
                    break;
                case SLPivotTableStyleTypeValues.Dark2:
                    this.Name = "PivotStyleDark2";
                    break;
                case SLPivotTableStyleTypeValues.Dark3:
                    this.Name = "PivotStyleDark3";
                    break;
                case SLPivotTableStyleTypeValues.Dark4:
                    this.Name = "PivotStyleDark4";
                    break;
                case SLPivotTableStyleTypeValues.Dark5:
                    this.Name = "PivotStyleDark5";
                    break;
                case SLPivotTableStyleTypeValues.Dark6:
                    this.Name = "PivotStyleDark6";
                    break;
                case SLPivotTableStyleTypeValues.Dark7:
                    this.Name = "PivotStyleDark7";
                    break;
                case SLPivotTableStyleTypeValues.Dark8:
                    this.Name = "PivotStyleDark8";
                    break;
                case SLPivotTableStyleTypeValues.Dark9:
                    this.Name = "PivotStyleDark9";
                    break;
                case SLPivotTableStyleTypeValues.Dark10:
                    this.Name = "PivotStyleDark10";
                    break;
                case SLPivotTableStyleTypeValues.Dark11:
                    this.Name = "PivotStyleDark11";
                    break;
                case SLPivotTableStyleTypeValues.Dark12:
                    this.Name = "PivotStyleDark12";
                    break;
                case SLPivotTableStyleTypeValues.Dark13:
                    this.Name = "PivotStyleDark13";
                    break;
                case SLPivotTableStyleTypeValues.Dark14:
                    this.Name = "PivotStyleDark14";
                    break;
                case SLPivotTableStyleTypeValues.Dark15:
                    this.Name = "PivotStyleDark15";
                    break;
                case SLPivotTableStyleTypeValues.Dark16:
                    this.Name = "PivotStyleDark16";
                    break;
                case SLPivotTableStyleTypeValues.Dark17:
                    this.Name = "PivotStyleDark17";
                    break;
                case SLPivotTableStyleTypeValues.Dark18:
                    this.Name = "PivotStyleDark18";
                    break;
                case SLPivotTableStyleTypeValues.Dark19:
                    this.Name = "PivotStyleDark19";
                    break;
                case SLPivotTableStyleTypeValues.Dark20:
                    this.Name = "PivotStyleDark20";
                    break;
                case SLPivotTableStyleTypeValues.Dark21:
                    this.Name = "PivotStyleDark21";
                    break;
                case SLPivotTableStyleTypeValues.Dark22:
                    this.Name = "PivotStyleDark22";
                    break;
                case SLPivotTableStyleTypeValues.Dark23:
                    this.Name = "PivotStyleDark23";
                    break;
                case SLPivotTableStyleTypeValues.Dark24:
                    this.Name = "PivotStyleDark24";
                    break;
                case SLPivotTableStyleTypeValues.Dark25:
                    this.Name = "PivotStyleDark25";
                    break;
                case SLPivotTableStyleTypeValues.Dark26:
                    this.Name = "PivotStyleDark26";
                    break;
                case SLPivotTableStyleTypeValues.Dark27:
                    this.Name = "PivotStyleDark27";
                    break;
                case SLPivotTableStyleTypeValues.Dark28:
                    this.Name = "PivotStyleDark28";
                    break;
            }
        }

        internal SLPivotTableStyle Clone()
        {
            SLPivotTableStyle pts = new SLPivotTableStyle();
            pts.Name = this.Name;
            pts.ShowRowHeaders = this.ShowRowHeaders;
            pts.ShowColumnHeaders = this.ShowColumnHeaders;
            pts.ShowRowStripes = this.ShowRowStripes;
            pts.ShowColumnStripes = this.ShowColumnStripes;
            pts.ShowLastColumn = this.ShowLastColumn;

            return pts;
        }
    }
}
