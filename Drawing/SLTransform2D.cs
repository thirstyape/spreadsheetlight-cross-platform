using System;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLight.Drawing
{
    internal class SLTransform2D
    {
        internal bool HasOffset;
        internal SLOffset Offset { get; set; }
        internal bool HasExtents;
        internal SLExtents Extents { get; set; }

        internal int? Rotation { get; set; }
        internal bool? HorizontalFlip { get; set; }
        internal bool? VerticalFlip { get; set; }

        internal SLTransform2D()
        {
            this.HasOffset = false;
            this.Offset = new SLOffset();
            this.HasExtents = false;
            this.Extents = new SLExtents();

            this.Rotation = null;
            this.HorizontalFlip = null;
            this.VerticalFlip = null;
        }

        internal A.Transform2D ToTransform2D()
        {
            A.Transform2D trans = new A.Transform2D();
            if (this.HasOffset) trans.Offset = this.Offset.ToOffset();
            if (this.HasExtents) trans.Extents = this.Extents.ToExtents();

            if (this.Rotation != null) trans.Rotation = this.Rotation.Value;
            if (this.HorizontalFlip != null) trans.HorizontalFlip = this.HorizontalFlip.Value;
            if (this.VerticalFlip != null) trans.VerticalFlip = this.VerticalFlip.Value;

            return trans;
        }

        internal SLTransform2D Clone()
        {
            SLTransform2D trans = new SLTransform2D();
            trans.HasOffset = this.HasOffset;
            trans.Offset = this.Offset.Clone();
            trans.HasExtents = this.HasExtents;
            trans.Extents = this.Extents.Clone();

            trans.Rotation = this.Rotation;
            trans.HorizontalFlip = this.HorizontalFlip;
            trans.VerticalFlip = this.VerticalFlip;

            return trans;
        }
    }
}
