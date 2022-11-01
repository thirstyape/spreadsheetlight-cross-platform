using System;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// For BubbleSize, Minus, Plus, Values, YValues
    /// </summary>
    internal class SLNumberDataSourceType
    {
        private bool bUseNumberReference;
        internal bool UseNumberReference
        {
            get { return bUseNumberReference; }
            set
            {
                bUseNumberReference = value;
                if (value)
                {
                    bUseNumberReference = true;
                    bUseNumberLiteral = false;
                }
            }
        }
        internal SLNumberReference NumberReference { get; set; }

        private bool bUseNumberLiteral;
        internal bool UseNumberLiteral
        {
            get { return bUseNumberLiteral; }
            set
            {
                bUseNumberLiteral = value;
                if (value)
                {
                    bUseNumberReference = false;
                    bUseNumberLiteral = true;
                }
            }
        }
        internal SLNumberLiteral NumberLiteral { get; set; }

        internal SLNumberDataSourceType()
        {
            this.UseNumberReference = true;

            this.NumberReference = new SLNumberReference();
            this.NumberLiteral = new SLNumberLiteral();
        }

        internal C.BubbleSize ToBubbleSize()
        {
            C.BubbleSize bs = new C.BubbleSize();
            if (UseNumberReference) bs.NumberReference = this.NumberReference.ToNumberReference();
            if (UseNumberLiteral) bs.NumberLiteral = this.NumberLiteral.ToNumberLiteral();

            return bs;
        }

        internal C.Minus ToMinus()
        {
            C.Minus minus = new C.Minus();
            if (UseNumberReference) minus.NumberReference = this.NumberReference.ToNumberReference();
            if (UseNumberLiteral) minus.NumberLiteral = this.NumberLiteral.ToNumberLiteral();

            return minus;
        }

        internal C.Plus ToPlus()
        {
            C.Plus plus = new C.Plus();
            if (UseNumberReference) plus.NumberReference = this.NumberReference.ToNumberReference();
            if (UseNumberLiteral) plus.NumberLiteral = this.NumberLiteral.ToNumberLiteral();

            return plus;
        }

        internal C.Values ToValues()
        {
            C.Values v = new C.Values();
            if (UseNumberReference) v.NumberReference = this.NumberReference.ToNumberReference();
            if (UseNumberLiteral) v.NumberLiteral = this.NumberLiteral.ToNumberLiteral();

            return v;
        }

        internal C.YValues ToYValues()
        {
            C.YValues yv = new C.YValues();
            if (UseNumberReference) yv.NumberReference = this.NumberReference.ToNumberReference();
            if (UseNumberLiteral) yv.NumberLiteral = this.NumberLiteral.ToNumberLiteral();

            return yv;
        }

        internal SLNumberDataSourceType Clone()
        {
            SLNumberDataSourceType ndst = new SLNumberDataSourceType();
            ndst.bUseNumberReference = this.bUseNumberReference;
            ndst.NumberReference = this.NumberReference.Clone();
            ndst.bUseNumberLiteral = this.bUseNumberLiteral;
            ndst.NumberLiteral = this.NumberLiteral.Clone();

            return ndst;
        }
    }
}
