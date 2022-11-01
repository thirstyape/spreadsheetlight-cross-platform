using System;
using System.Collections.Generic;
using System.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    internal class SLDataSeries
    {
        internal SLDataSeriesChartType ChartType { get; set; }

        internal uint Index { get; set; }
        internal uint Order { get; set; }

        // this is SeriesText
        internal bool? IsStringReference;
        internal SLStringReference StringReference { get; set; }
        internal string NumericValue { get; set; }

        internal SLDataSeriesOptions Options { get; set; }

        //PictureOptions

        internal Dictionary<int, SLDataPointOptions> DataPointOptionsList { get; set; }

        internal SLGroupDataLabelOptions GroupDataLabelOptions { get; set; }
        internal Dictionary<int, SLDataLabelOptions> DataLabelOptionsList { get; set; }

        //List<Trendline>
        //List<ErrorBars>

        //category
        //value

        //xval
        //yval

        internal SLNumberDataSourceType BubbleSize { get; set; }

        internal SLAxisDataSourceType AxisData { get; set; }
        internal SLNumberDataSourceType NumberData { get; set; }

        internal SLDataSeries(List<System.Drawing.Color> ThemeColors, bool ThrowExceptionsIfAny)
        {
            this.ChartType = SLDataSeriesChartType.None;

            this.Index = 0;
            this.Order = 0;

            this.IsStringReference = null;
            this.StringReference = new SLStringReference();
            this.NumericValue = string.Empty;

            this.Options = new SLDataSeriesOptions(ThemeColors, ThrowExceptionsIfAny);

            this.DataPointOptionsList = new Dictionary<int, SLDataPointOptions>();

            this.GroupDataLabelOptions = null;
            this.DataLabelOptionsList = new Dictionary<int, SLDataLabelOptions>();

            this.BubbleSize = new SLNumberDataSourceType();

            this.AxisData = new SLAxisDataSourceType();
            this.NumberData = new SLNumberDataSourceType();
        }

        internal C.PieChartSeries ToPieChartSeries(bool IsStylish)
        {
            C.PieChartSeries pcs = new C.PieChartSeries();
            pcs.Index = new C.Index() { Val = this.Index };
            pcs.Order = new C.Order() { Val = this.Order };

            if (this.IsStringReference != null)
            {
                pcs.SeriesText = new C.SeriesText();
                if (this.IsStringReference.Value)
                {
                    pcs.SeriesText.StringReference = this.StringReference.ToStringReference();
                }
                else
                {
                    pcs.SeriesText.NumericValue = new C.NumericValue(this.NumericValue);
                }
            }

            if (this.Options.ShapeProperties.HasShapeProperties)
            {
                pcs.ChartShapeProperties = this.Options.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            if (this.Options.iExplosion != null)
            {
                pcs.Explosion = new C.Explosion() { Val = this.Options.Explosion };
            }

            if (this.DataPointOptionsList.Count > 0)
            {
                List<int> indexlist = this.DataPointOptionsList.Keys.ToList<int>();
                indexlist.Sort();
                int index;
                for (int i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    pcs.Append(this.DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if (this.GroupDataLabelOptions != null || this.DataLabelOptionsList.Count > 0)
            {
                if (this.GroupDataLabelOptions == null)
                {
                    SLGroupDataLabelOptions gdloptions = new SLGroupDataLabelOptions(this.Options.ShapeProperties.listThemeColors, this.Options.ShapeProperties.ThrowExceptionsIfAny);
                    pcs.Append(gdloptions.ToDataLabels(this.DataLabelOptionsList, true));
                }
                else
                {
                    pcs.Append(this.GroupDataLabelOptions.ToDataLabels(this.DataLabelOptionsList, false));
                }
            }

            pcs.Append(this.AxisData.ToCategoryAxisData());
            pcs.Append(this.NumberData.ToValues());
            
            return pcs;
        }

        internal C.RadarChartSeries ToRadarChartSeries(bool IsStylish)
        {
            C.RadarChartSeries rcs = new C.RadarChartSeries();
            rcs.Index = new C.Index() { Val = this.Index };
            rcs.Order = new C.Order() { Val = this.Order };

            if (this.IsStringReference != null)
            {
                rcs.SeriesText = new C.SeriesText();
                if (this.IsStringReference.Value)
                {
                    rcs.SeriesText.StringReference = this.StringReference.ToStringReference();
                }
                else
                {
                    rcs.SeriesText.NumericValue = new C.NumericValue(this.NumericValue);
                }
            }

            if (this.Options.ShapeProperties.HasShapeProperties)
            {
                rcs.ChartShapeProperties = this.Options.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            if (this.Options.Marker.HasMarker)
            {
                rcs.Marker = this.Options.Marker.ToMarker(IsStylish);
            }

            if (this.DataPointOptionsList.Count > 0)
            {
                List<int> indexlist = this.DataPointOptionsList.Keys.ToList<int>();
                indexlist.Sort();
                int index;
                for (int i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    rcs.Append(this.DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if (this.GroupDataLabelOptions != null || this.DataLabelOptionsList.Count > 0)
            {
                if (this.GroupDataLabelOptions == null)
                {
                    SLGroupDataLabelOptions gdloptions = new SLGroupDataLabelOptions(this.Options.ShapeProperties.listThemeColors, this.Options.ShapeProperties.ThrowExceptionsIfAny);
                    rcs.Append(gdloptions.ToDataLabels(this.DataLabelOptionsList, true));
                }
                else
                {
                    rcs.Append(this.GroupDataLabelOptions.ToDataLabels(this.DataLabelOptionsList, false));
                }
            }

            rcs.Append(this.AxisData.ToCategoryAxisData());
            rcs.Append(this.NumberData.ToValues());

            return rcs;
        }

        internal C.AreaChartSeries ToAreaChartSeries(bool IsStylish)
        {
            C.AreaChartSeries acs = new C.AreaChartSeries();
            acs.Index = new C.Index() { Val = this.Index };
            acs.Order = new C.Order() { Val = this.Order };

            if (this.IsStringReference != null)
            {
                acs.SeriesText = new C.SeriesText();
                if (this.IsStringReference.Value)
                {
                    acs.SeriesText.StringReference = this.StringReference.ToStringReference();
                }
                else
                {
                    acs.SeriesText.NumericValue = new C.NumericValue(this.NumericValue);
                }
            }

            if (this.Options.ShapeProperties.HasShapeProperties)
            {
                acs.ChartShapeProperties = this.Options.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            //PictureOptions

            if (this.DataPointOptionsList.Count > 0)
            {
                List<int> indexlist = this.DataPointOptionsList.Keys.ToList<int>();
                indexlist.Sort();
                int index;
                for (int i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    acs.Append(this.DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if (this.GroupDataLabelOptions != null || this.DataLabelOptionsList.Count > 0)
            {
                if (this.GroupDataLabelOptions == null)
                {
                    SLGroupDataLabelOptions gdloptions = new SLGroupDataLabelOptions(this.Options.ShapeProperties.listThemeColors, this.Options.ShapeProperties.ThrowExceptionsIfAny);
                    acs.Append(gdloptions.ToDataLabels(this.DataLabelOptionsList, true));
                }
                else
                {
                    acs.Append(this.GroupDataLabelOptions.ToDataLabels(this.DataLabelOptionsList, false));
                }
            }

            acs.Append(this.AxisData.ToCategoryAxisData());
            acs.Append(this.NumberData.ToValues());

            return acs;
        }

        internal C.BarChartSeries ToBarChartSeries(bool IsStylish)
        {
            C.BarChartSeries bcs = new C.BarChartSeries();
            bcs.Index = new C.Index() { Val = this.Index };
            bcs.Order = new C.Order() { Val = this.Order };

            if (this.IsStringReference != null)
            {
                bcs.SeriesText = new C.SeriesText();
                if (this.IsStringReference.Value)
                {
                    bcs.SeriesText.StringReference = this.StringReference.ToStringReference();
                }
                else
                {
                    bcs.SeriesText.NumericValue = new C.NumericValue(this.NumericValue);
                }
            }

            if (this.Options.ShapeProperties.HasShapeProperties)
            {
                bcs.ChartShapeProperties = this.Options.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            bcs.InvertIfNegative = new C.InvertIfNegative() { Val = this.Options.InvertIfNegative ?? false };

            //PictureOptions

            if (this.DataPointOptionsList.Count > 0)
            {
                List<int> indexlist = this.DataPointOptionsList.Keys.ToList<int>();
                indexlist.Sort();
                int index;
                for (int i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    bcs.Append(this.DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if (this.GroupDataLabelOptions != null || this.DataLabelOptionsList.Count > 0)
            {
                if (this.GroupDataLabelOptions == null)
                {
                    SLGroupDataLabelOptions gdloptions = new SLGroupDataLabelOptions(this.Options.ShapeProperties.listThemeColors, this.Options.ShapeProperties.ThrowExceptionsIfAny);
                    bcs.Append(gdloptions.ToDataLabels(this.DataLabelOptionsList, true));
                }
                else
                {
                    bcs.Append(this.GroupDataLabelOptions.ToDataLabels(this.DataLabelOptionsList, false));
                }
            }

            bcs.Append(this.AxisData.ToCategoryAxisData());
            bcs.Append(this.NumberData.ToValues());

            if (this.Options.vShape != null)
            {
                bcs.Append(new C.Shape() { Val = this.Options.vShape.Value });
            }

            return bcs;
        }

        internal C.ScatterChartSeries ToScatterChartSeries(bool IsStylish)
        {
            C.ScatterChartSeries scs = new C.ScatterChartSeries();
            scs.Index = new C.Index() { Val = this.Index };
            scs.Order = new C.Order() { Val = this.Order };

            if (this.IsStringReference != null)
            {
                scs.SeriesText = new C.SeriesText();
                if (this.IsStringReference.Value)
                {
                    scs.SeriesText.StringReference = this.StringReference.ToStringReference();
                }
                else
                {
                    scs.SeriesText.NumericValue = new C.NumericValue(this.NumericValue);
                }
            }

            if (this.Options.ShapeProperties.HasShapeProperties)
            {
                scs.ChartShapeProperties = this.Options.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            if (this.Options.Marker.HasMarker)
            {
                scs.Marker = this.Options.Marker.ToMarker(IsStylish);
            }

            if (this.DataPointOptionsList.Count > 0)
            {
                List<int> indexlist = this.DataPointOptionsList.Keys.ToList<int>();
                indexlist.Sort();
                int index;
                for (int i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    scs.Append(this.DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if (this.GroupDataLabelOptions != null || this.DataLabelOptionsList.Count > 0)
            {
                if (this.GroupDataLabelOptions == null)
                {
                    SLGroupDataLabelOptions gdloptions = new SLGroupDataLabelOptions(this.Options.ShapeProperties.listThemeColors, this.Options.ShapeProperties.ThrowExceptionsIfAny);
                    scs.Append(gdloptions.ToDataLabels(this.DataLabelOptionsList, true));
                }
                else
                {
                    scs.Append(this.GroupDataLabelOptions.ToDataLabels(this.DataLabelOptionsList, false));
                }
            }

            scs.Append(this.AxisData.ToXValues());
            scs.Append(this.NumberData.ToYValues());

            scs.Append(new C.Smooth() { Val = this.Options.Smooth });

            return scs;
        }

        internal C.LineChartSeries ToLineChartSeries(bool IsStylish)
        {
            C.LineChartSeries lcs = new C.LineChartSeries();
            lcs.Index = new C.Index() { Val = this.Index };
            lcs.Order = new C.Order() { Val = this.Order };

            if (this.IsStringReference != null)
            {
                lcs.SeriesText = new C.SeriesText();
                if (this.IsStringReference.Value)
                {
                    lcs.SeriesText.StringReference = this.StringReference.ToStringReference();
                }
                else
                {
                    lcs.SeriesText.NumericValue = new C.NumericValue(this.NumericValue);
                }
            }

            if (this.Options.ShapeProperties.HasShapeProperties)
            {
                lcs.ChartShapeProperties = this.Options.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            if (this.Options.Marker.HasMarker)
            {
                lcs.Marker = this.Options.Marker.ToMarker(IsStylish);
            }

            if (this.DataPointOptionsList.Count > 0)
            {
                List<int> indexlist = this.DataPointOptionsList.Keys.ToList<int>();
                indexlist.Sort();
                int index;
                for (int i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    lcs.Append(this.DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if (this.GroupDataLabelOptions != null || this.DataLabelOptionsList.Count > 0)
            {
                if (this.GroupDataLabelOptions == null)
                {
                    SLGroupDataLabelOptions gdloptions = new SLGroupDataLabelOptions(this.Options.ShapeProperties.listThemeColors, this.Options.ShapeProperties.ThrowExceptionsIfAny);
                    lcs.Append(gdloptions.ToDataLabels(this.DataLabelOptionsList, true));
                }
                else
                {
                    lcs.Append(this.GroupDataLabelOptions.ToDataLabels(this.DataLabelOptionsList, false));
                }
            }

            lcs.Append(this.AxisData.ToCategoryAxisData());
            lcs.Append(this.NumberData.ToValues());

            lcs.Append(new C.Smooth() { Val = this.Options.Smooth });

            return lcs;
        }

        internal C.BubbleChartSeries ToBubbleChartSeries(bool IsStylish)
        {
            C.BubbleChartSeries bcs = new C.BubbleChartSeries();
            bcs.Index = new C.Index() { Val = this.Index };
            bcs.Order = new C.Order() { Val = this.Order };

            if (this.IsStringReference != null)
            {
                bcs.SeriesText = new C.SeriesText();
                if (this.IsStringReference.Value)
                {
                    bcs.SeriesText.StringReference = this.StringReference.ToStringReference();
                }
                else
                {
                    bcs.SeriesText.NumericValue = new C.NumericValue(this.NumericValue);
                }
            }

            if (this.Options.ShapeProperties.HasShapeProperties)
            {
                bcs.ChartShapeProperties = this.Options.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            bcs.InvertIfNegative = new C.InvertIfNegative() { Val = this.Options.InvertIfNegative ?? false };

            if (this.DataPointOptionsList.Count > 0)
            {
                List<int> indexlist = this.DataPointOptionsList.Keys.ToList<int>();
                indexlist.Sort();
                int index;
                for (int i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    bcs.Append(this.DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if (this.GroupDataLabelOptions != null || this.DataLabelOptionsList.Count > 0)
            {
                if (this.GroupDataLabelOptions == null)
                {
                    SLGroupDataLabelOptions gdloptions = new SLGroupDataLabelOptions(this.Options.ShapeProperties.listThemeColors, this.Options.ShapeProperties.ThrowExceptionsIfAny);
                    bcs.Append(gdloptions.ToDataLabels(this.DataLabelOptionsList, true));
                }
                else
                {
                    bcs.Append(this.GroupDataLabelOptions.ToDataLabels(this.DataLabelOptionsList, false));
                }
            }

            bcs.Append(this.AxisData.ToXValues());
            bcs.Append(this.NumberData.ToYValues());
            bcs.Append(this.BubbleSize.ToBubbleSize());

            if (this.Options.bBubble3D != null)
            {
                bcs.Append(new C.Bubble3D() { Val = this.Options.Bubble3D });
            }

            return bcs;
        }

        internal C.SurfaceChartSeries ToSurfaceChartSeries(bool IsStylish)
        {
            C.SurfaceChartSeries scs = new C.SurfaceChartSeries();
            scs.Index = new C.Index() { Val = this.Index };
            scs.Order = new C.Order() { Val = this.Order };

            if (this.IsStringReference != null)
            {
                scs.SeriesText = new C.SeriesText();
                if (this.IsStringReference.Value)
                {
                    scs.SeriesText.StringReference = this.StringReference.ToStringReference();
                }
                else
                {
                    scs.SeriesText.NumericValue = new C.NumericValue(this.NumericValue);
                }
            }

            if (this.Options.ShapeProperties.HasShapeProperties)
            {
                scs.ChartShapeProperties = this.Options.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            scs.Append(this.AxisData.ToCategoryAxisData());
            scs.Append(this.NumberData.ToValues());

            return scs;
        }

        internal SLDataSeries Clone()
        {
            SLDataSeries ds = new SLDataSeries(this.Options.ShapeProperties.listThemeColors, this.Options.ShapeProperties.ThrowExceptionsIfAny);
            ds.ChartType = this.ChartType;
            ds.Index = this.Index;
            ds.Order = this.Order;
            ds.IsStringReference = this.IsStringReference;
            ds.StringReference = this.StringReference.Clone();
            ds.NumericValue = this.NumericValue;
            ds.Options = this.Options.Clone();

            List<int> keys = this.DataPointOptionsList.Keys.ToList<int>();
            ds.DataPointOptionsList = new Dictionary<int, SLDataPointOptions>();
            foreach (int index in keys)
            {
                ds.DataPointOptionsList[index] = this.DataPointOptionsList[index].Clone();
            }

            if (this.GroupDataLabelOptions != null) ds.GroupDataLabelOptions = this.GroupDataLabelOptions.Clone();

            keys = this.DataLabelOptionsList.Keys.ToList<int>();
            ds.DataLabelOptionsList = new Dictionary<int, SLDataLabelOptions>();
            foreach (int index in keys)
            {
                ds.DataLabelOptionsList[index] = this.DataLabelOptionsList[index].Clone();
            }

            ds.BubbleSize = this.BubbleSize.Clone();
            ds.AxisData = this.AxisData.Clone();
            ds.NumberData = this.NumberData.Clone();

            return ds;
        }
    }
}
