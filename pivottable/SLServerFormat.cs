using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLServerFormat
    {
        internal string Culture { get; set; }
        internal string Format { get; set; }

        internal SLServerFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Culture = "";
            this.Format = "";
        }

        internal void FromServerFormat(ServerFormat sf)
        {
            this.SetAllNull();

            if (sf.Culture != null) this.Culture = sf.Culture.Value;
            if (sf.Format != null) this.Format = sf.Format.Value;
        }

        internal ServerFormat ToServerFormat()
        {
            ServerFormat sf = new ServerFormat();
            if (this.Culture != null && this.Culture.Length > 0) sf.Culture = this.Culture;
            if (this.Format != null && this.Format.Length > 0) sf.Format = this.Format;

            return sf;
        }

        internal SLServerFormat Clone()
        {
            SLServerFormat sf = new SLServerFormat();
            sf.Culture = this.Culture;
            sf.Format = this.Format;

            return sf;
        }
    }
}
