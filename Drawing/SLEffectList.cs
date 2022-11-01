using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Drawing
{
    /// <summary>
    /// Encapsulates properties and methods for specifying effects such as glow, shadows, reflection and soft edges.
    /// This simulates the DocumentFormat.OpenXml.Drawing.EffectList class.
    /// </summary>
    public class SLEffectList
    {
        internal List<System.Drawing.Color> listThemeColors;

        internal bool HasEffectList
        {
            get
            {
                return this.Glow.HasGlow || this.Shadow.IsInnerShadow != null
                    || this.Reflection.HasReflection || this.SoftEdge.HasSoftEdge;
            }
        }

        // A.Blur is not accessible from Excel! Don't know what values to allow...

        internal SLGlow Glow { get; set; }

        internal SLShadowEffect Shadow { get; set; }

        internal SLReflection Reflection { get; set; }

        internal SLSoftEdge SoftEdge { get; set; }

        internal SLEffectList(List<System.Drawing.Color> ThemeColors)
        {
            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }

            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Glow = new SLGlow(this.listThemeColors);
            this.Shadow = new SLShadowEffect(this.listThemeColors);
            this.Reflection = new SLReflection();
            this.SoftEdge = new SLSoftEdge();
        }

        internal A.EffectList ToEffectList()
        {
            A.EffectList el = new A.EffectList();

            if (this.Glow.HasGlow)
            {
                el.Glow = this.Glow.ToGlow();
            }

            if (this.Shadow.IsInnerShadow != null)
            {
                if (this.Shadow.IsInnerShadow.Value)
                {
                    el.InnerShadow = this.Shadow.ToInnerShadow();
                }
                else
                {
                    el.OuterShadow = this.Shadow.ToOuterShadow();
                }
            }

            if (this.Reflection.HasReflection)
            {
                el.Reflection = this.Reflection.ToReflection();
            }

            if (this.SoftEdge.HasSoftEdge)
            {
                el.SoftEdge = this.SoftEdge.ToSoftEdge();
            }

            return el;
        }

        internal SLEffectList Clone()
        {
            SLEffectList el = new SLEffectList(this.listThemeColors);
            el.Glow = this.Glow.Clone();
            el.Shadow = this.Shadow.Clone();
            el.Reflection = this.Reflection.Clone();
            el.SoftEdge = this.SoftEdge.Clone();

            return el;
        }
    }
}
