using System;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    /// <summary>
    /// Encapsulates properties and methods for specifying cell styles. This simulates the DocumentFormat.OpenXml.Spreadsheet.CellStyle class.
    /// </summary>
    public class SLCellStyle
    {
        /// <summary>
        /// Name of the cell style.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Specifies a zero-based index referencing a CellFormat in the CellStyleFormats class.
        /// </summary>
        public uint FormatId { get; set; }

        /// <summary>
        /// Specifies the index of a built-in cell style.
        /// </summary>
        public uint? BuiltinId { get; set; }

        /// <summary>
        /// Specifies that the formatting is for an outline style.
        /// </summary>
        public uint? OutlineLevel { get; set; }

        /// <summary>
        /// Specifies if the style is shown in the application user interface.
        /// </summary>
        public bool? Hidden { get; set; }

        /// <summary>
        /// Specifies if the built-in cell style is customized.
        /// </summary>
        public bool? CustomBuiltin { get; set; }

        /// <summary>
        /// Initializes an instance of SLCellStyle.
        /// </summary>
        public SLCellStyle()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Name = null;
            this.FormatId = 0;
            this.BuiltinId = null;
            this.OutlineLevel = null;
            this.Hidden = null;
            this.CustomBuiltin = null;
        }

        internal void FromCellStyle(CellStyle cs)
        {
            this.SetAllNull();

            if (cs.Name != null) this.Name = cs.Name.Value;
            if (cs.FormatId != null) this.FormatId = cs.FormatId.Value;
            if (cs.BuiltinId != null) this.BuiltinId = cs.BuiltinId.Value;
            if (cs.OutlineLevel != null) this.OutlineLevel = cs.OutlineLevel.Value;
            if (cs.Hidden != null) this.Hidden = cs.Hidden.Value;
            if (cs.CustomBuiltin != null) this.CustomBuiltin = cs.CustomBuiltin.Value;
        }

        internal CellStyle ToCellStyle()
        {
            CellStyle cs = new CellStyle();
            if (this.Name != null) cs.Name = this.Name;
            cs.FormatId = this.FormatId;
            if (this.BuiltinId != null) cs.BuiltinId = this.BuiltinId.Value;
            if (this.OutlineLevel != null) cs.OutlineLevel = this.OutlineLevel.Value;
            if (this.Hidden != null) cs.Hidden = this.Hidden.Value;
            if (this.CustomBuiltin != null) cs.CustomBuiltin = this.CustomBuiltin.Value;

            return cs;
        }

        internal void FromHash(string Hash)
        {
            this.SetAllNull();
            string[] sa = Hash.Split(new string[] { SLConstants.XmlCellStyleAttributeSeparator }, StringSplitOptions.None);

            if (sa.Length >= 6)
            {
                // weird if the actual name *is* "null"...
                if (!sa[0].Equals("null")) this.Name = sa[0];
                else this.Name = string.Empty;

                this.FormatId = uint.Parse(sa[1]);

                if (!sa[2].Equals("null")) this.BuiltinId = uint.Parse(sa[2]);

                if (!sa[3].Equals("null")) this.OutlineLevel = uint.Parse(sa[3]);

                if (!sa[4].Equals("null"))
                {
                    if (sa[4].Equals("true")) this.Hidden = true;
                    else if (sa[4].Equals("false")) this.Hidden = false;
                }

                if (!sa[5].Equals("null"))
                {
                    if (sa[5].Equals("true")) this.CustomBuiltin = true;
                    else if (sa[5].Equals("false")) this.CustomBuiltin = false;
                }
            }
        }

        internal string ToHash()
        {
            StringBuilder sb = new StringBuilder();

            if (this.Name != null) sb.AppendFormat("{0}{1}", this.Name, SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            sb.AppendFormat("{0}{1}", this.FormatId, SLConstants.XmlCellStyleAttributeSeparator);

            if (this.BuiltinId != null) sb.AppendFormat("{0}{1}", this.BuiltinId.Value, SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            if (this.OutlineLevel != null) sb.AppendFormat("{0}{1}", this.OutlineLevel.Value, SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            if (this.Hidden != null) sb.AppendFormat("{0}{1}", (this.Hidden.Value) ? "true" : "false", SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            if (this.CustomBuiltin != null) sb.AppendFormat("{0}{1}", (this.CustomBuiltin.Value) ? "true" : "false", SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            return sb.ToString();
        }

        internal string WriteToXmlTag()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<x:cellStyle");
            if (this.Name != null) sb.AppendFormat(" name=\"{0}\"", this.Name);
            sb.AppendFormat(" xfId=\"{0}\"", this.FormatId);
            if (this.BuiltinId != null) sb.AppendFormat(" builtinId=\"{0}\"", this.BuiltinId.Value);
            if (this.OutlineLevel != null) sb.AppendFormat(" iLevel=\"{0}\"", this.OutlineLevel.Value);
            if (this.Hidden != null) sb.AppendFormat(" hidden=\"{0}\"", this.Hidden.Value);
            if (this.CustomBuiltin != null) sb.AppendFormat(" customBuiltin=\"{0}\"", this.CustomBuiltin.Value);
            sb.Append(" />");

            return sb.ToString();
        }

        internal SLCellStyle Clone()
        {
            SLCellStyle cs = new SLCellStyle();
            cs.Name = this.Name;
            cs.FormatId = this.FormatId;
            cs.BuiltinId = this.BuiltinId;
            cs.OutlineLevel = this.OutlineLevel;
            cs.Hidden = this.Hidden;
            cs.CustomBuiltin = this.CustomBuiltin;

            return cs;
        }
    }
}
