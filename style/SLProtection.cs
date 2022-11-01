using System;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for cell content protection. The properties don't take effect unless the worksheet is protected. This simulates the DocumentFormat.OpenXml.Spreadsheet.Protection class.
    /// </summary>
    public class SLProtection
    {
        /// <summary>
        /// Specifies if the cell is locked. If locked and the worksheet is protected, then the worksheet's protection options are ignored.
        /// </summary>
        public bool? Locked { get; set; }

        /// <summary>
        /// Specifies if the cell is hidden. If hidden and the worksheet is protected, then cell contents are hidden and only cell values are displayed. For example, the cell formula is hidden, but the value of the cell formula is still displayed.
        /// </summary>
        public bool? Hidden { get; set; }

        /// <summary>
        /// Initializes an instance of SLProtection.
        /// </summary>
        public SLProtection()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Locked = null;
            this.Hidden = null;
        }

        internal void FromProtection(Protection p)
        {
            this.SetAllNull();

            if (p.Locked != null)
            {
                this.Locked = p.Locked.Value;
            }

            if (p.Hidden != null)
            {
                this.Hidden = p.Hidden.Value;
            }
        }

        internal Protection ToProtection()
        {
            Protection p = new Protection();
            if (this.Locked != null) p.Locked = this.Locked.Value;
            if (this.Hidden != null) p.Hidden = this.Hidden.Value;

            return p;
        }

        internal void FromHash(string Hash)
        {
            this.SetAllNull();

            string[] sa = Hash.Split(new string[] { SLConstants.XmlProtectionAttributeSeparator }, StringSplitOptions.None);
            if (sa.Length >= 2)
            {
                if (!sa[0].Equals("null")) this.Locked = bool.Parse(sa[0]);

                if (!sa[1].Equals("null")) this.Hidden = bool.Parse(sa[1]);
            }
        }

        internal string ToHash()
        {
            StringBuilder sb = new StringBuilder();

            if (this.Locked != null) sb.AppendFormat("{0}{1}", this.Locked.Value, SLConstants.XmlProtectionAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlProtectionAttributeSeparator);

            if (this.Hidden != null) sb.AppendFormat("{0}{1}", this.Hidden.Value, SLConstants.XmlProtectionAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlProtectionAttributeSeparator);

            return sb.ToString();
        }

        internal string WriteToXmlTag()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<x:protection");
            if (this.Locked != null) sb.AppendFormat(" locked=\"{0}\"", this.Locked.Value ? "1" : "0");
            if (this.Hidden != null) sb.AppendFormat(" hidden=\"{0}\"", this.Hidden.Value ? "1" : "0");
            sb.Append(" />");

            return sb.ToString();
        }

        internal SLProtection Clone()
        {
            SLProtection p = new SLProtection();
            p.Locked = this.Locked;
            p.Hidden = this.Hidden;

            return p;
        }
    }
}
