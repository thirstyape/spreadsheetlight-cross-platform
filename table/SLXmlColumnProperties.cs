using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLXmlColumnProperties
    {
        internal uint MapId { get; set; }
        internal string XPath { get; set; }
        internal bool? Denormalized { get; set; }
        internal XmlDataValues XmlDataType { get; set; }

        internal SLXmlColumnProperties()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.MapId = 0;
            this.XPath = string.Empty;
            this.Denormalized = null;
            this.XmlDataType = XmlDataValues.AnyType;
        }

        internal void FromXmlColumnProperties(XmlColumnProperties xcp)
        {
            this.SetAllNull();

            if (xcp.MapId != null) this.MapId = xcp.MapId.Value;
            if (xcp.XPath != null) this.XPath = xcp.XPath.Value;
            if (xcp.Denormalized != null && xcp.Denormalized.Value) this.Denormalized = xcp.Denormalized.Value;
            if (xcp.XmlDataType != null) this.XmlDataType = xcp.XmlDataType.Value;
        }

        internal XmlColumnProperties ToXmlColumnProperties()
        {
            XmlColumnProperties xcp = new XmlColumnProperties();
            xcp.MapId = this.MapId;
            xcp.XPath = this.XPath;
            if (this.Denormalized != null && this.Denormalized.Value) xcp.Denormalized = this.Denormalized.Value;
            xcp.XmlDataType = this.XmlDataType;

            return xcp;
        }

        internal SLXmlColumnProperties Clone()
        {
            SLXmlColumnProperties xcp = new SLXmlColumnProperties();
            xcp.MapId = this.MapId;
            xcp.XPath = this.XPath;
            xcp.Denormalized = this.Denormalized;
            xcp.XmlDataType = this.XmlDataType;

            return xcp;
        }
    }
}
