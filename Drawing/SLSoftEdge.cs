using System;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Drawing
{
    /// <summary>
    /// Encapsulates properties and methods for specifying soft edges.
    /// This simulates the DocumentFormat.OpenXml.Drawing.SoftEdge class.
    /// </summary>
    public class SLSoftEdge
    {
        internal bool HasSoftEdge;

        internal decimal decRadius;
        // Probably an example of the marketing team, usability/testing team, technical documentation
        // team and the development team not talking to each other properly.
        // "Normal" people relate to the word "Size". Technical people have no problems with "Radius".
        // I'm gonna go with the technical documentation slash development team here.
        // Also, the Open XML specs use radius. Ahh... but the tech doc people wrote the specs...
        /// <summary>
        /// Also known as "Size", and is measured in points. The suggested range is 0 pt to 100 pt (both inclusive).
        /// </summary>
        public decimal Radius
        {
            get { return decRadius; }
            set
            {
                HasSoftEdge = true;
                decRadius = value;
                if (decRadius < 0m) decRadius = 0m;
                if (decRadius > 2147483647m) decRadius = 2147483647m;
            }
        }

        /// <summary>
        /// Set no soft edge.
        /// </summary>
        public void SetNoSoftEdge()
        {
            this.HasSoftEdge = false;
            this.decRadius = 0;
        }

        internal SLSoftEdge()
        {
            this.HasSoftEdge = false;
            this.decRadius = 0;
        }

        internal A.SoftEdge ToSoftEdge()
        {
            A.SoftEdge se = new A.SoftEdge();
            se.Radius = SLA.SLDrawingTool.CalculatePositiveCoordinate(decRadius);

            return se;
        }

        internal SLSoftEdge Clone()
        {
            SLSoftEdge se = new SLSoftEdge();
            se.HasSoftEdge = this.HasSoftEdge;
            se.decRadius = this.decRadius;

            return se;
        }
    }
}
