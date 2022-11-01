using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLNumberingFormat
    {
        private uint iNumberFormatId;
        internal uint NumberFormatId
        {
            get { return iNumberFormatId; }
            set { iNumberFormatId = value; }
        }

        private string sFormatCode;
        internal string FormatCode
        {
            get { return sFormatCode; }
            set { sFormatCode = value; }
        }

        internal SLNumberingFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.iNumberFormatId = 0;
            this.sFormatCode = string.Empty;
        }

        internal void FromNumberingFormat(NumberingFormat nf)
        {
            this.SetAllNull();

            if (nf.NumberFormatId != null)
            {
                this.NumberFormatId = nf.NumberFormatId.Value;
            }
            else
            {
                this.NumberFormatId = 0;
            }

            if (nf.FormatCode != null)
            {
                this.FormatCode = nf.FormatCode.Value;
            }
            else
            {
                this.FormatCode = string.Empty;
            }
        }

        internal NumberingFormat ToNumberingFormat()
        {
            NumberingFormat nf = new NumberingFormat();
            nf.NumberFormatId = this.NumberFormatId;
            nf.FormatCode = this.FormatCode;

            return nf;
        }

        internal void FromHash(string Hash)
        {
            this.FormatCode = Hash;
        }

        internal string ToHash()
        {
            return this.FormatCode;
        }

        internal SLNumberingFormat Clone()
        {
            SLNumberingFormat nf = new SLNumberingFormat();
            nf.iNumberFormatId = this.iNumberFormatId;
            nf.sFormatCode = this.sFormatCode;

            return nf;
        }
    }
}
