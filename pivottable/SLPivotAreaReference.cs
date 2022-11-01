using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLPivotAreaReference
    {
        internal List<uint> FieldItems { get; set; }

        internal uint? Field { get; set; }
        //internal uint Count { get; set; }
        internal bool Selected { get; set; }
        internal bool ByPosition { get; set; }
        internal bool Relative { get; set; }
        internal bool DefaultSubtotal { get; set; }
        internal bool SumSubtotal { get; set; }
        internal bool CountASubtotal { get; set; }
        internal bool AverageSubtotal { get; set; }
        internal bool MaxSubtotal { get; set; }
        internal bool MinSubtotal { get; set; }
        internal bool ApplyProductInSubtotal { get; set; }
        internal bool CountSubtotal { get; set; }
        internal bool ApplyStandardDeviationInSubtotal { get; set; }
        internal bool ApplyStandardDeviationPInSubtotal { get; set; }
        internal bool ApplyVarianceInSubtotal { get; set; }
        internal bool ApplyVariancePInSubtotal { get; set; }

        internal SLPivotAreaReference()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.FieldItems = new List<uint>();
            this.Field = null;
            //this.Count = 0;
            this.Selected = true;
            this.ByPosition = false;
            this.Relative = false;
            this.DefaultSubtotal = false;
            this.SumSubtotal = false;
            this.CountASubtotal = false;
            this.AverageSubtotal = false;
            this.MaxSubtotal = false;
            this.MinSubtotal = false;
            this.ApplyProductInSubtotal = false;
            this.CountSubtotal = false;
            this.ApplyStandardDeviationInSubtotal = false;
            this.ApplyStandardDeviationPInSubtotal = false;
            this.ApplyVarianceInSubtotal = false;
            this.ApplyVariancePInSubtotal = false;
        }

        internal void FromPivotAreaReference(PivotAreaReference par)
        {
            this.SetAllNull();

            if (par.Field != null) this.Field = par.Field.Value;
            if (par.Selected != null) this.Selected = par.Selected.Value;
            if (par.ByPosition != null) this.ByPosition = par.ByPosition.Value;
            if (par.Relative != null) this.Relative = par.Relative.Value;
            if (par.DefaultSubtotal != null) this.DefaultSubtotal = par.DefaultSubtotal.Value;
            if (par.SumSubtotal != null) this.SumSubtotal = par.SumSubtotal.Value;
            if (par.CountASubtotal != null) this.CountASubtotal = par.CountASubtotal.Value;
            if (par.AverageSubtotal != null) this.AverageSubtotal = par.AverageSubtotal.Value;
            if (par.MaxSubtotal != null) this.MaxSubtotal = par.MaxSubtotal.Value;
            if (par.MinSubtotal != null) this.MinSubtotal = par.MinSubtotal.Value;
            if (par.ApplyProductInSubtotal != null) this.ApplyProductInSubtotal = par.ApplyProductInSubtotal.Value;
            if (par.CountSubtotal != null) this.CountSubtotal = par.CountSubtotal.Value;
            if (par.ApplyStandardDeviationInSubtotal != null) this.ApplyStandardDeviationInSubtotal = par.ApplyStandardDeviationInSubtotal.Value;
            if (par.ApplyStandardDeviationPInSubtotal != null) this.ApplyStandardDeviationPInSubtotal = par.ApplyStandardDeviationPInSubtotal.Value;
            if (par.ApplyVarianceInSubtotal != null) this.ApplyVarianceInSubtotal = par.ApplyVarianceInSubtotal.Value;
            if (par.ApplyVariancePInSubtotal != null) this.ApplyVariancePInSubtotal = par.ApplyVariancePInSubtotal.Value;

            FieldItem fi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(par))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(FieldItem))
                    {
                        fi = (FieldItem)oxr.LoadCurrentElement();
                        // the Val property is required
                        this.FieldItems.Add(fi.Val.Value);
                    }
                }
            }
        }

        internal PivotAreaReference ToPivotAreaReference()
        {
            PivotAreaReference par = new PivotAreaReference();
            if (this.Field != null) par.Field = this.Field.Value;
            par.Count = (uint)this.FieldItems.Count;
            if (this.Selected != true) par.Selected = this.Selected;
            if (this.ByPosition != false) par.ByPosition = this.ByPosition;
            if (this.Relative != false) par.Relative = this.Relative;
            if (this.DefaultSubtotal != false) par.DefaultSubtotal = this.DefaultSubtotal;
            if (this.SumSubtotal != false) par.SumSubtotal = this.SumSubtotal;
            if (this.CountASubtotal != false) par.CountASubtotal = this.CountASubtotal;
            if (this.AverageSubtotal != false) par.AverageSubtotal = this.AverageSubtotal;
            if (this.MaxSubtotal != false) par.MaxSubtotal = this.MaxSubtotal;
            if (this.MinSubtotal != false) par.MinSubtotal = this.MinSubtotal;
            if (this.ApplyProductInSubtotal != false) par.ApplyProductInSubtotal = this.ApplyProductInSubtotal;
            if (this.CountSubtotal != false) par.CountSubtotal = this.CountSubtotal;
            if (this.ApplyStandardDeviationInSubtotal != false) par.ApplyStandardDeviationInSubtotal = this.ApplyStandardDeviationInSubtotal;
            if (this.ApplyStandardDeviationPInSubtotal != false) par.ApplyStandardDeviationPInSubtotal = this.ApplyStandardDeviationPInSubtotal;
            if (this.ApplyVarianceInSubtotal != false) par.ApplyVarianceInSubtotal = this.ApplyVarianceInSubtotal;
            if (this.ApplyVariancePInSubtotal != false) par.ApplyVariancePInSubtotal = this.ApplyVariancePInSubtotal;

            foreach (uint i in this.FieldItems)
            {
                par.Append(new FieldItem() { Val = i });
            }

            return par;
        }

        internal SLPivotAreaReference Clone()
        {
            SLPivotAreaReference par = new SLPivotAreaReference();
            par.Field = this.Field;
            par.Selected = this.Selected;
            par.ByPosition = this.ByPosition;
            par.Relative = this.Relative;
            par.DefaultSubtotal = this.DefaultSubtotal;
            par.SumSubtotal = this.SumSubtotal;
            par.CountASubtotal = this.CountASubtotal;
            par.AverageSubtotal = this.AverageSubtotal;
            par.MaxSubtotal = this.MaxSubtotal;
            par.MinSubtotal = this.MinSubtotal;
            par.ApplyProductInSubtotal = this.ApplyProductInSubtotal;
            par.CountSubtotal = this.CountSubtotal;
            par.ApplyStandardDeviationInSubtotal = this.ApplyStandardDeviationInSubtotal;
            par.ApplyStandardDeviationPInSubtotal = this.ApplyStandardDeviationPInSubtotal;
            par.ApplyVarianceInSubtotal = this.ApplyVarianceInSubtotal;
            par.ApplyVariancePInSubtotal = this.ApplyVariancePInSubtotal;

            par.FieldItems = new List<uint>();
            foreach (uint i in this.FieldItems)
            {
                par.FieldItems.Add(i);
            }

            return par;
        }
    }
}
