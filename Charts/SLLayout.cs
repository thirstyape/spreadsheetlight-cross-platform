using System;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    internal class SLLayout
    {
        internal C.LayoutTargetValues LayoutTarget { get; set; }
        internal C.LayoutModeValues LeftMode { get; set; }
        internal C.LayoutModeValues TopMode { get; set; }
        internal C.LayoutModeValues WidthMode { get; set; }
        internal C.LayoutModeValues HeightMode { get; set; }
        internal double? Left { get; set; }
        internal double? Top { get; set; }
        internal double? Width { get; set; }
        internal double? Height { get; set; }

        internal SLLayout()
        {
            this.LayoutTarget = C.LayoutTargetValues.Outer;
            this.LeftMode = C.LayoutModeValues.Factor;
            this.TopMode = C.LayoutModeValues.Factor;
            this.WidthMode = C.LayoutModeValues.Factor;
            this.HeightMode = C.LayoutModeValues.Factor;
            this.Left = null;
            this.Top = null;
            this.Width = null;
            this.Height = null;
        }

        internal C.Layout ToLayout()
        {
            C.Layout layout = new C.Layout();

            if (this.LayoutTarget != C.LayoutTargetValues.Outer
                || this.LeftMode != C.LayoutModeValues.Factor || this.TopMode != C.LayoutModeValues.Factor
                || this.WidthMode != C.LayoutModeValues.Factor || this.HeightMode != C.LayoutModeValues.Factor
                || this.Left != null || this.Top != null || this.Width != null || this.Height != null)
            {
                layout.ManualLayout = new C.ManualLayout();
                if (this.LayoutTarget != C.LayoutTargetValues.Outer) layout.ManualLayout.LayoutTarget = new C.LayoutTarget() { Val = this.LayoutTarget };
                if (this.LeftMode != C.LayoutModeValues.Factor) layout.ManualLayout.LeftMode = new C.LeftMode() { Val = this.LeftMode };
                if (this.TopMode != C.LayoutModeValues.Factor) layout.ManualLayout.TopMode = new C.TopMode() { Val = this.TopMode };
                if (this.WidthMode != C.LayoutModeValues.Factor) layout.ManualLayout.WidthMode = new C.WidthMode() { Val = this.WidthMode };
                if (this.HeightMode != C.LayoutModeValues.Factor) layout.ManualLayout.HeightMode = new C.HeightMode() { Val = this.HeightMode };
                if (this.Left != null) layout.ManualLayout.Left = new C.Left() { Val = this.Left.Value };
                if (this.Top != null) layout.ManualLayout.Top = new C.Top() { Val = this.Top.Value };
                if (this.Width != null) layout.ManualLayout.Width = new C.Width() { Val = this.Width.Value };
                if (this.Height != null) layout.ManualLayout.Height = new C.Height() { Val = this.Height.Value };
            }

            return layout;
        }

        internal SLLayout Clone()
        {
            SLLayout l = new SLLayout();
            l.LayoutTarget = this.LayoutTarget;
            l.LeftMode = this.LeftMode;
            l.TopMode = this.TopMode;
            l.WidthMode = this.WidthMode;
            l.HeightMode = this.HeightMode;
            l.Left = this.Left;
            l.Top = this.Top;
            l.Width = this.Width;
            l.Height = this.Height;

            return l;
        }
    }
}
