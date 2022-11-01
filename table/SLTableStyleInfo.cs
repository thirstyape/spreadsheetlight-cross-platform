using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Specifies the built-in table style type.
    /// </summary>
    public enum SLTableStyleTypeValues
    {
        /// <summary>
        /// Table Style Light 1
        /// </summary>
        Light1 = 0,
        /// <summary>
        /// Table Style Light 2
        /// </summary>
        Light2,
        /// <summary>
        /// Table Style Light 3
        /// </summary>
        Light3,
        /// <summary>
        /// Table Style Light 4
        /// </summary>
        Light4,
        /// <summary>
        /// Table Style Light 5
        /// </summary>
        Light5,
        /// <summary>
        /// Table Style Light 6
        /// </summary>
        Light6,
        /// <summary>
        /// Table Style Light 7
        /// </summary>
        Light7,
        /// <summary>
        /// Table Style Light 8
        /// </summary>
        Light8,
        /// <summary>
        /// Table Style Light 9
        /// </summary>
        Light9,
        /// <summary>
        /// Table Style Light 10
        /// </summary>
        Light10,
        /// <summary>
        /// Table Style Light 11
        /// </summary>
        Light11,
        /// <summary>
        /// Table Style Light 12
        /// </summary>
        Light12,
        /// <summary>
        /// Table Style Light 13
        /// </summary>
        Light13,
        /// <summary>
        /// Table Style Light 14
        /// </summary>
        Light14,
        /// <summary>
        /// Table Style Light 15
        /// </summary>
        Light15,
        /// <summary>
        /// Table Style Light 16
        /// </summary>
        Light16,
        /// <summary>
        /// Table Style Light 17
        /// </summary>
        Light17,
        /// <summary>
        /// Table Style Light 18
        /// </summary>
        Light18,
        /// <summary>
        /// Table Style Light 19
        /// </summary>
        Light19,
        /// <summary>
        /// Table Style Light 20
        /// </summary>
        Light20,
        /// <summary>
        /// Table Style Light 21
        /// </summary>
        Light21,
        /// <summary>
        /// Table Style Medium 1
        /// </summary>
        Medium1,
        /// <summary>
        /// Table Style Medium 2
        /// </summary>
        Medium2,
        /// <summary>
        /// Table Style Medium 3
        /// </summary>
        Medium3,
        /// <summary>
        /// Table Style Medium 4
        /// </summary>
        Medium4,
        /// <summary>
        /// Table Style Medium 5
        /// </summary>
        Medium5,
        /// <summary>
        /// Table Style Medium 6
        /// </summary>
        Medium6,
        /// <summary>
        /// Table Style Medium 7
        /// </summary>
        Medium7,
        /// <summary>
        /// Table Style Medium 8
        /// </summary>
        Medium8,
        /// <summary>
        /// Table Style Medium 9
        /// </summary>
        Medium9,
        /// <summary>
        /// Table Style Medium 10
        /// </summary>
        Medium10,
        /// <summary>
        /// Table Style Medium 11
        /// </summary>
        Medium11,
        /// <summary>
        /// Table Style Medium 12
        /// </summary>
        Medium12,
        /// <summary>
        /// Table Style Medium 13
        /// </summary>
        Medium13,
        /// <summary>
        /// Table Style Medium 14
        /// </summary>
        Medium14,
        /// <summary>
        /// Table Style Medium 15
        /// </summary>
        Medium15,
        /// <summary>
        /// Table Style Medium 16
        /// </summary>
        Medium16,
        /// <summary>
        /// Table Style Medium 17
        /// </summary>
        Medium17,
        /// <summary>
        /// Table Style Medium 18
        /// </summary>
        Medium18,
        /// <summary>
        /// Table Style Medium 19
        /// </summary>
        Medium19,
        /// <summary>
        /// Table Style Medium 20
        /// </summary>
        Medium20,
        /// <summary>
        /// Table Style Medium 21
        /// </summary>
        Medium21,
        /// <summary>
        /// Table Style Medium 22
        /// </summary>
        Medium22,
        /// <summary>
        /// Table Style Medium 23
        /// </summary>
        Medium23,
        /// <summary>
        /// Table Style Medium 24
        /// </summary>
        Medium24,
        /// <summary>
        /// Table Style Medium 25
        /// </summary>
        Medium25,
        /// <summary>
        /// Table Style Medium 26
        /// </summary>
        Medium26,
        /// <summary>
        /// Table Style Medium 27
        /// </summary>
        Medium27,
        /// <summary>
        /// Table Style Medium 28
        /// </summary>
        Medium28,
        /// <summary>
        /// Table Style Dark 1
        /// </summary>
        Dark1,
        /// <summary>
        /// Table Style Dark 2
        /// </summary>
        Dark2,
        /// <summary>
        /// Table Style Dark 3
        /// </summary>
        Dark3,
        /// <summary>
        /// Table Style Dark 4
        /// </summary>
        Dark4,
        /// <summary>
        /// Table Style Dark 5
        /// </summary>
        Dark5,
        /// <summary>
        /// Table Style Dark 6
        /// </summary>
        Dark6,
        /// <summary>
        /// Table Style Dark 7
        /// </summary>
        Dark7,
        /// <summary>
        /// Table Style Dark 8
        /// </summary>
        Dark8,
        /// <summary>
        /// Table Style Dark 9
        /// </summary>
        Dark9,
        /// <summary>
        /// Table Style Dark 10
        /// </summary>
        Dark10,
        /// <summary>
        /// Table Style Dark 11
        /// </summary>
        Dark11
    }

    internal class SLTableStyleInfo
    {
        internal string Name { get; set; }
        internal bool? ShowFirstColumn { get; set; }
        internal bool? ShowLastColumn { get; set; }
        internal bool? ShowRowStripes { get; set; }
        internal bool? ShowColumnStripes { get; set; }

        internal SLTableStyleInfo()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Name = null;
            this.ShowFirstColumn = null;
            this.ShowLastColumn = null;
            this.ShowRowStripes = null;
            this.ShowColumnStripes = null;
        }

        internal void FromTableStyleInfo(TableStyleInfo tsi)
        {
            this.SetAllNull();

            if (tsi.Name != null) this.Name = tsi.Name.Value;
            if (tsi.ShowFirstColumn != null) this.ShowFirstColumn = tsi.ShowFirstColumn.Value;
            if (tsi.ShowLastColumn != null) this.ShowLastColumn = tsi.ShowLastColumn.Value;
            if (tsi.ShowRowStripes != null) this.ShowRowStripes = tsi.ShowRowStripes.Value;
            if (tsi.ShowColumnStripes != null) this.ShowColumnStripes = tsi.ShowColumnStripes.Value;
        }

        internal TableStyleInfo ToTableStyleInfo()
        {
            TableStyleInfo tsi = new TableStyleInfo();
            if (this.Name != null) tsi.Name = this.Name;
            if (this.ShowFirstColumn != null) tsi.ShowFirstColumn = this.ShowFirstColumn.Value;
            if (this.ShowLastColumn != null) tsi.ShowLastColumn = this.ShowLastColumn.Value;
            if (this.ShowRowStripes != null) tsi.ShowRowStripes = this.ShowRowStripes.Value;
            if (this.ShowColumnStripes != null) tsi.ShowColumnStripes = this.ShowColumnStripes.Value;

            return tsi;
        }

        internal void SetTableStyle(SLTableStyleTypeValues tstyle)
        {
            switch (tstyle)
            {
                case SLTableStyleTypeValues.Light1:
                    this.Name = "TableStyleLight1";
                    break;
                case SLTableStyleTypeValues.Light2:
                    this.Name = "TableStyleLight2";
                    break;
                case SLTableStyleTypeValues.Light3:
                    this.Name = "TableStyleLight3";
                    break;
                case SLTableStyleTypeValues.Light4:
                    this.Name = "TableStyleLight4";
                    break;
                case SLTableStyleTypeValues.Light5:
                    this.Name = "TableStyleLight5";
                    break;
                case SLTableStyleTypeValues.Light6:
                    this.Name = "TableStyleLight6";
                    break;
                case SLTableStyleTypeValues.Light7:
                    this.Name = "TableStyleLight7";
                    break;
                case SLTableStyleTypeValues.Light8:
                    this.Name = "TableStyleLight8";
                    break;
                case SLTableStyleTypeValues.Light9:
                    this.Name = "TableStyleLight9";
                    break;
                case SLTableStyleTypeValues.Light10:
                    this.Name = "TableStyleLight10";
                    break;
                case SLTableStyleTypeValues.Light11:
                    this.Name = "TableStyleLight11";
                    break;
                case SLTableStyleTypeValues.Light12:
                    this.Name = "TableStyleLight12";
                    break;
                case SLTableStyleTypeValues.Light13:
                    this.Name = "TableStyleLight13";
                    break;
                case SLTableStyleTypeValues.Light14:
                    this.Name = "TableStyleLight14";
                    break;
                case SLTableStyleTypeValues.Light15:
                    this.Name = "TableStyleLight15";
                    break;
                case SLTableStyleTypeValues.Light16:
                    this.Name = "TableStyleLight16";
                    break;
                case SLTableStyleTypeValues.Light17:
                    this.Name = "TableStyleLight17";
                    break;
                case SLTableStyleTypeValues.Light18:
                    this.Name = "TableStyleLight18";
                    break;
                case SLTableStyleTypeValues.Light19:
                    this.Name = "TableStyleLight19";
                    break;
                case SLTableStyleTypeValues.Light20:
                    this.Name = "TableStyleLight20";
                    break;
                case SLTableStyleTypeValues.Light21:
                    this.Name = "TableStyleLight21";
                    break;
                case SLTableStyleTypeValues.Medium1:
                    this.Name = "TableStyleMedium1";
                    break;
                case SLTableStyleTypeValues.Medium2:
                    this.Name = "TableStyleMedium2";
                    break;
                case SLTableStyleTypeValues.Medium3:
                    this.Name = "TableStyleMedium3";
                    break;
                case SLTableStyleTypeValues.Medium4:
                    this.Name = "TableStyleMedium4";
                    break;
                case SLTableStyleTypeValues.Medium5:
                    this.Name = "TableStyleMedium5";
                    break;
                case SLTableStyleTypeValues.Medium6:
                    this.Name = "TableStyleMedium6";
                    break;
                case SLTableStyleTypeValues.Medium7:
                    this.Name = "TableStyleMedium7";
                    break;
                case SLTableStyleTypeValues.Medium8:
                    this.Name = "TableStyleMedium8";
                    break;
                case SLTableStyleTypeValues.Medium9:
                    this.Name = "TableStyleMedium9";
                    break;
                case SLTableStyleTypeValues.Medium10:
                    this.Name = "TableStyleMedium10";
                    break;
                case SLTableStyleTypeValues.Medium11:
                    this.Name = "TableStyleMedium11";
                    break;
                case SLTableStyleTypeValues.Medium12:
                    this.Name = "TableStyleMedium12";
                    break;
                case SLTableStyleTypeValues.Medium13:
                    this.Name = "TableStyleMedium13";
                    break;
                case SLTableStyleTypeValues.Medium14:
                    this.Name = "TableStyleMedium14";
                    break;
                case SLTableStyleTypeValues.Medium15:
                    this.Name = "TableStyleMedium15";
                    break;
                case SLTableStyleTypeValues.Medium16:
                    this.Name = "TableStyleMedium16";
                    break;
                case SLTableStyleTypeValues.Medium17:
                    this.Name = "TableStyleMedium17";
                    break;
                case SLTableStyleTypeValues.Medium18:
                    this.Name = "TableStyleMedium18";
                    break;
                case SLTableStyleTypeValues.Medium19:
                    this.Name = "TableStyleMedium19";
                    break;
                case SLTableStyleTypeValues.Medium20:
                    this.Name = "TableStyleMedium20";
                    break;
                case SLTableStyleTypeValues.Medium21:
                    this.Name = "TableStyleMedium21";
                    break;
                case SLTableStyleTypeValues.Medium22:
                    this.Name = "TableStyleMedium22";
                    break;
                case SLTableStyleTypeValues.Medium23:
                    this.Name = "TableStyleMedium23";
                    break;
                case SLTableStyleTypeValues.Medium24:
                    this.Name = "TableStyleMedium24";
                    break;
                case SLTableStyleTypeValues.Medium25:
                    this.Name = "TableStyleMedium25";
                    break;
                case SLTableStyleTypeValues.Medium26:
                    this.Name = "TableStyleMedium26";
                    break;
                case SLTableStyleTypeValues.Medium27:
                    this.Name = "TableStyleMedium27";
                    break;
                case SLTableStyleTypeValues.Medium28:
                    this.Name = "TableStyleMedium28";
                    break;
                case SLTableStyleTypeValues.Dark1:
                    this.Name = "TableStyleDark1";
                    break;
                case SLTableStyleTypeValues.Dark2:
                    this.Name = "TableStyleDark2";
                    break;
                case SLTableStyleTypeValues.Dark3:
                    this.Name = "TableStyleDark3";
                    break;
                case SLTableStyleTypeValues.Dark4:
                    this.Name = "TableStyleDark4";
                    break;
                case SLTableStyleTypeValues.Dark5:
                    this.Name = "TableStyleDark5";
                    break;
                case SLTableStyleTypeValues.Dark6:
                    this.Name = "TableStyleDark6";
                    break;
                case SLTableStyleTypeValues.Dark7:
                    this.Name = "TableStyleDark7";
                    break;
                case SLTableStyleTypeValues.Dark8:
                    this.Name = "TableStyleDark8";
                    break;
                case SLTableStyleTypeValues.Dark9:
                    this.Name = "TableStyleDark9";
                    break;
                case SLTableStyleTypeValues.Dark10:
                    this.Name = "TableStyleDark10";
                    break;
                case SLTableStyleTypeValues.Dark11:
                    this.Name = "TableStyleDark11";
                    break;
            }
        }

        internal SLTableStyleInfo Clone()
        {
            SLTableStyleInfo tsi = new SLTableStyleInfo();
            tsi.Name = this.Name;
            tsi.ShowFirstColumn = this.ShowFirstColumn;
            tsi.ShowLastColumn = this.ShowLastColumn;
            tsi.ShowRowStripes = this.ShowRowStripes;
            tsi.ShowColumnStripes = this.ShowColumnStripes;

            return tsi;
        }
    }
}
