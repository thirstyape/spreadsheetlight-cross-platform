using System;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLTableStyle
    {
        internal string TableStyleInnerXml;

        internal string Name;

        internal bool? Pivot { get; set; }

        internal bool? Table { get; set; }

        internal uint? Count { get; set; }

        internal SLTableStyle()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.TableStyleInnerXml = string.Empty;
            this.Name = string.Empty;
            this.Pivot = null;
            this.Table = null;
            this.Count = null;
        }

        internal void FromTableStyle(TableStyle ts)
        {
            this.SetAllNull();

            this.TableStyleInnerXml = ts.InnerXml;

            // this is a required field, so it can't be null, but just in case...
            if (ts.Name != null) this.Name = ts.Name.Value;
            else this.Name = string.Empty;

            if (ts.Pivot != null)
            {
                this.Pivot = ts.Pivot.Value;
            }

            if (ts.Table != null)
            {
                this.Table = ts.Table.Value;
            }

            if (ts.Count != null)
            {
                this.Count = ts.Count.Value;
            }
        }

        internal TableStyle ToTableStyle()
        {
            TableStyle ts = new TableStyle();
            ts.InnerXml = SLTool.RemoveNamespaceDeclaration(this.TableStyleInnerXml);
            ts.Name = this.Name;

            if (this.Pivot != null) ts.Pivot = this.Pivot.Value;
            if (this.Table != null) ts.Table = this.Table.Value;
            if (this.Count != null) ts.Count = this.Count.Value;

            return ts;
        }

        internal void FromHash(string Hash)
        {
            TableStyle ts = new TableStyle();

            string[] saElementAttribute = Hash.Split(new string[] { SLConstants.XmlTableStyleElementAttributeSeparator }, StringSplitOptions.None);

            if (saElementAttribute.Length >= 2)
            {
                ts.InnerXml = saElementAttribute[0];
                string[] sa = saElementAttribute[1].Split(new string[] { SLConstants.XmlTableStyleAttributeSeparator }, StringSplitOptions.None);
                if (sa.Length >= 4)
                {
                    ts.Name = sa[0];

                    if (!sa[1].Equals("null")) ts.Pivot = bool.Parse(sa[1]);

                    if (!sa[2].Equals("null")) ts.Table = bool.Parse(sa[2]);

                    if (!sa[3].Equals("null")) ts.Count = uint.Parse(sa[3]);
                }
            }

            this.FromTableStyle(ts);
        }

        internal string ToHash()
        {
            TableStyle ts = this.ToTableStyle();
            string sXml = SLTool.RemoveNamespaceDeclaration(ts.InnerXml);

            StringBuilder sb = new StringBuilder();

            sb.AppendFormat("{0}{1}", sXml, SLConstants.XmlTableStyleElementAttributeSeparator);

            sb.AppendFormat("{0}{1}", ts.Name.Value, SLConstants.XmlTableStyleAttributeSeparator);

            if (ts.Pivot != null) sb.AppendFormat("{0}{1}", ts.Pivot.Value, SLConstants.XmlTableStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlTableStyleAttributeSeparator);

            if (ts.Table != null) sb.AppendFormat("{0}{1}", ts.Table.Value, SLConstants.XmlTableStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlTableStyleAttributeSeparator);

            if (ts.Count != null) sb.AppendFormat("{0}{1}", ts.Count.Value, SLConstants.XmlTableStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlTableStyleAttributeSeparator);

            return sb.ToString();
        }

        internal string WriteToXmlTag()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("<x:tableStyle name=\"{0}\"", this.Name);
            if (this.Pivot != null && !this.Pivot.Value) sb.Append(" pivot=\"0\"");
            if (this.Table != null && !this.Table.Value) sb.Append(" table=\"0\"");
            if (this.Count != null) sb.AppendFormat(" count=\"{0}\"", this.Count.Value);

            if (this.TableStyleInnerXml.Length > 0)
            {
                sb.Append(">");
                sb.Append(this.TableStyleInnerXml);
                sb.Append("</x:tableStyle>");
            }
            else
            {
                sb.Append(" />");
            }

            return sb.ToString();
        }

        internal SLTableStyle Clone()
        {
            SLTableStyle ts = new SLTableStyle();
            ts.TableStyleInnerXml = this.TableStyleInnerXml;
            ts.Name = this.Name;
            ts.Pivot = this.Pivot;
            ts.Table = this.Table;
            ts.Count = this.Count;

            return ts;
        }
    }
}
