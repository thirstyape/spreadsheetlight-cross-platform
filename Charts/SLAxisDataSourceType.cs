using System;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// For CategoryAxisData and XValues
    /// </summary>
    internal class SLAxisDataSourceType
    {
        protected bool bUseMultiLevelStringReference;
        internal bool UseMultiLevelStringReference
        {
            get { return bUseMultiLevelStringReference; }
            set
            {
                bUseMultiLevelStringReference = value;
                if (value)
                {
                    bUseMultiLevelStringReference = true;
                    bUseNumberReference = false;
                    bUseNumberLiteral = false;
                    bUseStringReference = false;
                    bUseStringLiteral = false;
                }
            }
        }
        internal SLMultiLevelStringReference MultiLevelStringReference { get; set; }

        protected bool bUseNumberReference;
        internal bool UseNumberReference
        {
            get { return bUseNumberReference; }
            set
            {
                bUseNumberReference = value;
                if (value)
                {
                    bUseMultiLevelStringReference = false;
                    bUseNumberReference = true;
                    bUseNumberLiteral = false;
                    bUseStringReference = false;
                    bUseStringLiteral = false;
                }
            }
        }
        internal SLNumberReference NumberReference { get; set; }

        protected bool bUseNumberLiteral;
        internal bool UseNumberLiteral
        {
            get { return bUseNumberLiteral; }
            set
            {
                bUseNumberLiteral = value;
                if (value)
                {
                    bUseMultiLevelStringReference = false;
                    bUseNumberReference = false;
                    bUseNumberLiteral = true;
                    bUseStringReference = false;
                    bUseStringLiteral = false;
                }
            }
        }
        internal SLNumberLiteral NumberLiteral { get; set; }

        protected bool bUseStringReference;
        internal bool UseStringReference
        {
            get { return bUseStringReference; }
            set
            {
                bUseStringReference = value;
                if (value)
                {
                    bUseMultiLevelStringReference = false;
                    bUseNumberReference = false;
                    bUseNumberLiteral = false;
                    bUseStringReference = true;
                    bUseStringLiteral = false;
                }
            }
        }
        internal SLStringReference StringReference { get; set; }

        protected bool bUseStringLiteral;
        internal bool UseStringLiteral
        {
            get { return bUseStringLiteral; }
            set
            {
                bUseStringLiteral = value;
                if (value)
                {
                    bUseMultiLevelStringReference = false;
                    bUseNumberReference = false;
                    bUseNumberLiteral = false;
                    bUseStringReference = false;
                    bUseStringLiteral = true;
                }
            }
        }
        internal SLStringLiteral StringLiteral { get; set; }

        internal SLAxisDataSourceType()
        {
            this.UseStringReference = true;

            this.MultiLevelStringReference = new SLMultiLevelStringReference();
            this.NumberReference = new SLNumberReference();
            this.NumberLiteral = new SLNumberLiteral();
            this.StringReference = new SLStringReference();
            this.StringLiteral = new SLStringLiteral();
        }

        internal C.CategoryAxisData ToCategoryAxisData()
        {
            C.CategoryAxisData cad = new C.CategoryAxisData();
            if (UseMultiLevelStringReference) cad.MultiLevelStringReference = this.MultiLevelStringReference.ToMultiLevelStringReference();
            if (UseNumberReference) cad.NumberReference = this.NumberReference.ToNumberReference();
            if (UseNumberLiteral) cad.NumberLiteral = this.NumberLiteral.ToNumberLiteral();
            if (UseStringReference) cad.StringReference = this.StringReference.ToStringReference();
            if (UseStringLiteral) cad.StringLiteral = this.StringLiteral.ToStringLiteral();

            return cad;
        }

        internal C.XValues ToXValues()
        {
            C.XValues xv = new C.XValues();
            if (UseMultiLevelStringReference) xv.MultiLevelStringReference = this.MultiLevelStringReference.ToMultiLevelStringReference();
            if (UseNumberReference) xv.NumberReference = this.NumberReference.ToNumberReference();
            if (UseNumberLiteral) xv.NumberLiteral = this.NumberLiteral.ToNumberLiteral();
            if (UseStringReference) xv.StringReference = this.StringReference.ToStringReference();
            if (UseStringLiteral) xv.StringLiteral = this.StringLiteral.ToStringLiteral();

            return xv;
        }

        internal SLAxisDataSourceType Clone()
        {
            SLAxisDataSourceType adst = new SLAxisDataSourceType();
            adst.bUseMultiLevelStringReference = this.bUseMultiLevelStringReference;
            adst.bUseNumberLiteral = this.bUseNumberLiteral;
            adst.bUseNumberReference = this.bUseNumberReference;
            adst.bUseStringLiteral = this.bUseStringLiteral;
            adst.bUseStringReference = this.bUseStringReference;

            adst.MultiLevelStringReference = this.MultiLevelStringReference.Clone();
            adst.NumberLiteral = this.NumberLiteral.Clone();
            adst.NumberReference = this.NumberReference.Clone();
            adst.StringLiteral = this.StringLiteral.Clone();
            adst.StringReference = this.StringReference.Clone();

            return adst;
        }
    }
}
