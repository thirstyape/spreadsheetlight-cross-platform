using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    internal enum SLDataSeriesChartType
    {
        // make sure to start from 0 because this is also going to be used in array indices
        DoughnutChart = 0,
        OfPieChartBar,
        OfPieChartPie,
        PieChart,
        RadarChartPrimary,
        RadarChartSecondary,
        AreaChartPrimary,
        AreaChartSecondary,
        BarChartColumnPrimary,
        BarChartColumnSecondary,
        BarChartBarPrimary,
        BarChartBarSecondary,
        ScatterChartPrimary,
        ScatterChartSecondary,
        LineChartPrimary,
        LineChartSecondary,
        // the following supposedly can't be used in combination charts
        Area3DChart,
        Bar3DChart,
        BubbleChart,
        Line3DChart,
        Pie3DChart,
        SurfaceChart,
        Surface3DChart,
        StockChart,
        // just for default purposes. Shouldn't affect memory or performance just because there's one more enumeration.
        None
    }

    /// <summary>
    /// Encapsulates properties and methods for setting plot areas in charts.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.PlotArea class.
    /// </summary>
    public class SLPlotArea
    {
        internal SLInternalChartType InternalChartType { get; set; }

        internal bool[] UsedChartTypes;
        internal SLChartOptions[] UsedChartOptions;
        internal List<SLDataSeries> DataSeries;

        internal SLLayout Layout { get; set; }

        internal SLTextAxis PrimaryTextAxis { get; set; }
        internal SLValueAxis PrimaryValueAxis { get; set; }
        internal SLSeriesAxis DepthAxis { get; set; }
        internal SLTextAxis SecondaryTextAxis { get; set; }
        internal SLValueAxis SecondaryValueAxis { get; set; }

        internal bool HasPrimaryAxes { get; set; }
        internal bool HasDepthAxis { get; set; }
        internal bool HasSecondaryAxes { get; set; }

        internal bool ShowDataTable { get; set; }
        internal SLDataTable DataTable { get; set; }

        internal SLA.SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Fill properties.
        /// </summary>
        public SLA.SLFill Fill { get { return this.ShapeProperties.Fill; } }

        /// <summary>
        /// Border properties.
        /// </summary>
        public SLA.SLLinePropertiesType Border { get { return this.ShapeProperties.Outline; } }

        /// <summary>
        /// Shadow properties.
        /// </summary>
        public SLA.SLShadowEffect Shadow { get { return this.ShapeProperties.EffectList.Shadow; } }

        /// <summary>
        /// Glow properties.
        /// </summary>
        public SLA.SLGlow Glow { get { return this.ShapeProperties.EffectList.Glow; } }

        /// <summary>
        /// Soft edge properties.
        /// </summary>
        public SLA.SLSoftEdge SoftEdge { get { return this.ShapeProperties.EffectList.SoftEdge; } }

        /// <summary>
        /// 3D format properties.
        /// </summary>
        public SLA.SLFormat3D Format3D { get { return this.ShapeProperties.Format3D; } }

        internal SLPlotArea(List<System.Drawing.Color> ThemeColors, bool Date1904, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.InternalChartType = SLInternalChartType.Bar;

            int NumberOfChartTypes = Enum.GetNames(typeof(SLDataSeriesChartType)).Length;
            this.UsedChartTypes = new bool[NumberOfChartTypes];
            this.UsedChartOptions = new SLChartOptions[NumberOfChartTypes];
            for (int i = 0; i < NumberOfChartTypes; ++i)
            {
                this.UsedChartTypes[i] = false;
                this.UsedChartOptions[i] = new SLChartOptions(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            }
            this.DataSeries = new List<SLDataSeries>();

            this.Layout = new SLLayout();

            this.PrimaryTextAxis = new SLTextAxis(ThemeColors, Date1904, IsStylish, ThrowExceptionsIfAny);
            this.PrimaryValueAxis = new SLValueAxis(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.DepthAxis = new SLSeriesAxis(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.SecondaryTextAxis = new SLTextAxis(ThemeColors, Date1904, IsStylish, ThrowExceptionsIfAny);
            this.SecondaryValueAxis = new SLValueAxis(ThemeColors, IsStylish, ThrowExceptionsIfAny);

            this.HasPrimaryAxes = false;
            this.HasDepthAxis = false;
            this.HasSecondaryAxes = false;

            this.ShowDataTable = false;
            this.DataTable = new SLDataTable(ThemeColors, IsStylish, ThrowExceptionsIfAny);

            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
            if (IsStylish)
            {
                this.ShapeProperties.Fill.SetNoFill();
                this.ShapeProperties.Outline.SetNoLine();
            }
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
        }

        internal void SetDataSeriesChartType(SLDataSeriesChartType ChartType)
        {
            for (int i = 0; i < this.DataSeries.Count; ++i)
            {
                this.DataSeries[i].ChartType = ChartType;
            }
        }

        internal void SetDataSeriesAutoAxisType()
        {
            // the first data series is good enough. In fact, AxisData should be identical for all.
            if (this.DataSeries.Count > 0)
            {
                // if it's a NumberReference, it might be a date
                if (this.DataSeries[0].AxisData.UseNumberReference)
                {
                    string sFormatCode = this.DataSeries[0].AxisData.NumberReference.NumberingCache.FormatCode;
                    if (SLTool.CheckIfFormatCodeIsDateRelated(sFormatCode))
                    {
                        this.PrimaryTextAxis.AxisType = SLAxisType.Date;
                        this.PrimaryTextAxis.FormatCode = sFormatCode;
                        this.PrimaryTextAxis.BaseUnit = C.TimeUnitValues.Days;
                        this.SecondaryTextAxis.AxisType = SLAxisType.Date;
                        this.SecondaryTextAxis.FormatCode = sFormatCode;
                        this.SecondaryTextAxis.BaseUnit = C.TimeUnitValues.Days;
                    }
                }
            }
        }

        internal C.PlotArea ToPlotArea(bool IsStylish)
        {
            C.PlotArea pa = new C.PlotArea();
            pa.Append(this.Layout.ToLayout());

            int iChartType;
            int i;

            // TODO: the rendering order is sort of listed in the following.
            // But apparently if you plot data series for doughnut first before bar-of-pie
            // it's different than if you plot bar-of-pie then doughnut.
            // Find out the "correct" order next version I suppose...

            // Excel 2010 apparently sets this by default for any chart...
            SLGroupDataLabelOptions gdlo = new SLGroupDataLabelOptions(this.ShapeProperties.listThemeColors, this.ShapeProperties.ThrowExceptionsIfAny);
            gdlo.ShowLegendKey = false;
            gdlo.ShowValue = false;
            gdlo.ShowCategoryName = false;
            gdlo.ShowSeriesName = false;
            gdlo.ShowPercentage = false;
            gdlo.ShowBubbleSize = false;

            #region Doughnut
            iChartType = (int)SLDataSeriesChartType.DoughnutChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.DoughnutChart dc = new C.DoughnutChart();
                dc.VaryColors = new C.VaryColors() { Val = this.UsedChartOptions[iChartType].VaryColors ?? true };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        dc.Append(this.DataSeries[i].ToPieChartSeries(IsStylish));
                    }
                }

                dc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                dc.Append(new C.FirstSliceAngle() { Val = this.UsedChartOptions[iChartType].FirstSliceAngle });
                dc.Append(new C.HoleSize() { Val = this.UsedChartOptions[iChartType].HoleSize });
                
                pa.Append(dc);
            }
            #endregion

            #region Bar-of-pie
            iChartType = (int)SLDataSeriesChartType.OfPieChartBar;
            if (this.UsedChartTypes[iChartType])
            {
                C.OfPieChart opc = new C.OfPieChart();
                opc.OfPieType = new C.OfPieType() { Val = C.OfPieValues.Bar };
                opc.VaryColors = new C.VaryColors() { Val = this.UsedChartOptions[iChartType].VaryColors ?? true };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        opc.Append(this.DataSeries[i].ToPieChartSeries(IsStylish));
                    }
                }

                opc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                opc.Append(new C.GapWidth() { Val = this.UsedChartOptions[iChartType].GapWidth });
                
                if (this.UsedChartOptions[iChartType].HasSplit)
                {
                    opc.Append(new C.SplitType() { Val = this.UsedChartOptions[iChartType].SplitType });
                    if (this.UsedChartOptions[iChartType].SplitType != C.SplitValues.Custom)
                    {
                        opc.Append(new C.SplitPosition() { Val = this.UsedChartOptions[iChartType].SplitPosition });
                    }
                    else
                    {
                        C.CustomSplit custsplit = new C.CustomSplit();
                        foreach (int iPiePoint in this.UsedChartOptions[iChartType].SecondPiePoints)
                        {
                            custsplit.Append(new C.SecondPiePoint() { Val = (uint)iPiePoint });
                        }
                        opc.Append(custsplit);
                    }
                }

                opc.Append(new C.SecondPieSize() { Val = this.UsedChartOptions[iChartType].SecondPieSize });

                if (this.UsedChartOptions[iChartType].SeriesLinesShapeProperties.HasShapeProperties)
                {
                    opc.Append(new C.SeriesLines()
                    {
                        ChartShapeProperties = this.UsedChartOptions[iChartType].SeriesLinesShapeProperties.ToChartShapeProperties(IsStylish)
                    });
                }
                else
                {
                    opc.Append(new C.SeriesLines());
                }

                pa.Append(opc);
            }
            #endregion

            #region Pie-of-pie
            iChartType = (int)SLDataSeriesChartType.OfPieChartPie;
            if (this.UsedChartTypes[iChartType])
            {
                C.OfPieChart opc = new C.OfPieChart();
                opc.OfPieType = new C.OfPieType() { Val = C.OfPieValues.Pie };
                opc.VaryColors = new C.VaryColors() { Val = this.UsedChartOptions[iChartType].VaryColors ?? true };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        opc.Append(this.DataSeries[i].ToPieChartSeries(IsStylish));
                    }
                }

                opc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                opc.Append(new C.GapWidth() { Val = this.UsedChartOptions[iChartType].GapWidth });

                if (this.UsedChartOptions[iChartType].HasSplit)
                {
                    opc.Append(new C.SplitType() { Val = this.UsedChartOptions[iChartType].SplitType });
                    if (this.UsedChartOptions[iChartType].SplitType != C.SplitValues.Custom)
                    {
                        opc.Append(new C.SplitPosition() { Val = this.UsedChartOptions[iChartType].SplitPosition });
                    }
                    else
                    {
                        C.CustomSplit custsplit = new C.CustomSplit();
                        foreach (int iPiePoint in this.UsedChartOptions[iChartType].SecondPiePoints)
                        {
                            custsplit.Append(new C.SecondPiePoint() { Val = (uint)iPiePoint });
                        }
                        opc.Append(custsplit);
                    }
                }

                opc.Append(new C.SecondPieSize() { Val = this.UsedChartOptions[iChartType].SecondPieSize });

                if (this.UsedChartOptions[iChartType].SeriesLinesShapeProperties.HasShapeProperties)
                {
                    opc.Append(new C.SeriesLines()
                    {
                        ChartShapeProperties = this.UsedChartOptions[iChartType].SeriesLinesShapeProperties.ToChartShapeProperties(IsStylish)
                    });
                }
                else
                {
                    opc.Append(new C.SeriesLines());
                }

                pa.Append(opc);
            }
            #endregion

            #region Pie
            iChartType = (int)SLDataSeriesChartType.PieChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.PieChart pc = new C.PieChart();
                pc.VaryColors = new C.VaryColors() { Val = this.UsedChartOptions[iChartType].VaryColors ?? true };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        pc.Append(this.DataSeries[i].ToPieChartSeries(IsStylish));
                    }
                }

                pc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                pc.Append(new C.FirstSliceAngle() { Val = this.UsedChartOptions[iChartType].FirstSliceAngle });

                pa.Append(pc);
            }
            #endregion

            #region Radar primary
            iChartType = (int)SLDataSeriesChartType.RadarChartPrimary;
            if (this.UsedChartTypes[iChartType])
            {
                C.RadarChart rc = new C.RadarChart();
                rc.RadarStyle = new C.RadarStyle() { Val = this.UsedChartOptions[iChartType].RadarStyle };
                rc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        rc.Append(this.DataSeries[i].ToRadarChartSeries(IsStylish));
                    }
                }

                rc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                rc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                rc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });

                pa.Append(rc);
            }
            #endregion

            #region Radar secondary
            iChartType = (int)SLDataSeriesChartType.RadarChartSecondary;
            if (this.UsedChartTypes[iChartType])
            {
                C.RadarChart rc = new C.RadarChart();
                rc.RadarStyle = new C.RadarStyle() { Val = this.UsedChartOptions[iChartType].RadarStyle };
                rc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        rc.Append(this.DataSeries[i].ToRadarChartSeries(IsStylish));
                    }
                }

                rc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                rc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis1 });
                rc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis2 });

                pa.Append(rc);
            }
            #endregion

            #region Area primary
            iChartType = (int)SLDataSeriesChartType.AreaChartPrimary;
            if (this.UsedChartTypes[iChartType])
            {
                C.AreaChart ac = new C.AreaChart();
                ac.Grouping = new C.Grouping() { Val = this.UsedChartOptions[iChartType].Grouping };
                ac.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        ac.Append(this.DataSeries[i].ToAreaChartSeries(IsStylish));
                    }
                }

                ac.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (this.UsedChartOptions[iChartType].HasDropLines)
                {
                    ac.Append(this.UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));
                }

                ac.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                ac.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });

                pa.Append(ac);
            }
            #endregion

            #region Area secondary
            iChartType = (int)SLDataSeriesChartType.AreaChartSecondary;
            if (this.UsedChartTypes[iChartType])
            {
                C.AreaChart ac = new C.AreaChart();
                ac.Grouping = new C.Grouping() { Val = this.UsedChartOptions[iChartType].Grouping };
                ac.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        ac.Append(this.DataSeries[i].ToAreaChartSeries(IsStylish));
                    }
                }

                ac.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (this.UsedChartOptions[iChartType].HasDropLines)
                {
                    ac.Append(this.UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));
                }

                ac.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis1 });
                ac.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis2 });

                pa.Append(ac);
            }
            #endregion

            #region Column primary
            iChartType = (int)SLDataSeriesChartType.BarChartColumnPrimary;
            if (this.UsedChartTypes[iChartType])
            {
                C.BarChart bc = new C.BarChart();
                bc.BarDirection = new C.BarDirection() { Val = C.BarDirectionValues.Column };
                bc.BarGrouping = new C.BarGrouping() { Val = this.UsedChartOptions[iChartType].BarGrouping };
                bc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        bc.Append(this.DataSeries[i].ToBarChartSeries(IsStylish));
                    }
                }

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth() { Val = this.UsedChartOptions[iChartType].GapWidth });

                if (this.UsedChartOptions[iChartType].Overlap != 0)
                {
                    bc.Append(new C.Overlap() { Val = this.UsedChartOptions[iChartType].Overlap });
                }

                bc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                bc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });

                pa.Append(bc);
            }
            #endregion

            #region Column secondary
            iChartType = (int)SLDataSeriesChartType.BarChartColumnSecondary;
            if (this.UsedChartTypes[iChartType])
            {
                C.BarChart bc = new C.BarChart();
                bc.BarDirection = new C.BarDirection() { Val = C.BarDirectionValues.Column };
                bc.BarGrouping = new C.BarGrouping() { Val = this.UsedChartOptions[iChartType].BarGrouping };
                bc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        bc.Append(this.DataSeries[i].ToBarChartSeries(IsStylish));
                    }
                }

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth() { Val = this.UsedChartOptions[iChartType].GapWidth });

                if (this.UsedChartOptions[iChartType].Overlap != 0)
                {
                    bc.Append(new C.Overlap() { Val = this.UsedChartOptions[iChartType].Overlap });
                }

                bc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis1 });
                bc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis2 });

                pa.Append(bc);
            }
            #endregion

            #region Bar primary
            iChartType = (int)SLDataSeriesChartType.BarChartBarPrimary;
            if (this.UsedChartTypes[iChartType])
            {
                C.BarChart bc = new C.BarChart();
                bc.BarDirection = new C.BarDirection() { Val = C.BarDirectionValues.Bar };
                bc.BarGrouping = new C.BarGrouping() { Val = this.UsedChartOptions[iChartType].BarGrouping };
                bc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        bc.Append(this.DataSeries[i].ToBarChartSeries(IsStylish));
                    }
                }

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth() { Val = this.UsedChartOptions[iChartType].GapWidth });

                if (this.UsedChartOptions[iChartType].Overlap != 0)
                {
                    bc.Append(new C.Overlap() { Val = this.UsedChartOptions[iChartType].Overlap });
                }

                bc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                bc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });

                pa.Append(bc);
            }
            #endregion

            #region Bar secondary
            iChartType = (int)SLDataSeriesChartType.BarChartBarSecondary;
            if (this.UsedChartTypes[iChartType])
            {
                C.BarChart bc = new C.BarChart();
                bc.BarDirection = new C.BarDirection() { Val = C.BarDirectionValues.Bar };
                bc.BarGrouping = new C.BarGrouping() { Val = this.UsedChartOptions[iChartType].BarGrouping };
                bc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        bc.Append(this.DataSeries[i].ToBarChartSeries(IsStylish));
                    }
                }

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth() { Val = this.UsedChartOptions[iChartType].GapWidth });

                if (this.UsedChartOptions[iChartType].Overlap != 0)
                {
                    bc.Append(new C.Overlap() { Val = this.UsedChartOptions[iChartType].Overlap });
                }

                bc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis1 });
                bc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis2 });

                pa.Append(bc);
            }
            #endregion

            #region Scatter primary
            iChartType = (int)SLDataSeriesChartType.ScatterChartPrimary;
            if (this.UsedChartTypes[iChartType])
            {
                C.ScatterChart sc = new C.ScatterChart();
                sc.ScatterStyle = new C.ScatterStyle() { Val = this.UsedChartOptions[iChartType].ScatterStyle };
                sc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        sc.Append(this.DataSeries[i].ToScatterChartSeries(IsStylish));
                    }
                }

                sc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });

                pa.Append(sc);
            }
            #endregion

            #region Scatter secondary
            iChartType = (int)SLDataSeriesChartType.ScatterChartSecondary;
            if (this.UsedChartTypes[iChartType])
            {
                C.ScatterChart sc = new C.ScatterChart();
                sc.ScatterStyle = new C.ScatterStyle() { Val = this.UsedChartOptions[iChartType].ScatterStyle };
                sc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        sc.Append(this.DataSeries[i].ToScatterChartSeries(IsStylish));
                    }
                }

                sc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                sc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis1 });
                sc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis2 });

                pa.Append(sc);
            }
            #endregion

            #region Line primary
            iChartType = (int)SLDataSeriesChartType.LineChartPrimary;
            if (this.UsedChartTypes[iChartType])
            {
                C.LineChart lc = new C.LineChart();
                lc.Grouping = new C.Grouping() { Val = this.UsedChartOptions[iChartType].Grouping };
                lc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        lc.Append(this.DataSeries[i].ToLineChartSeries(IsStylish));
                    }
                }

                lc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (this.UsedChartOptions[iChartType].HasDropLines)
                {
                    lc.Append(this.UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));
                }

                lc.Append(new C.ShowMarker() { Val = this.UsedChartOptions[iChartType].ShowMarker });
                lc.Append(new C.Smooth() { Val = this.UsedChartOptions[iChartType].Smooth });

                lc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                lc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });

                pa.Append(lc);
            }
            #endregion

            #region Line secondary
            iChartType = (int)SLDataSeriesChartType.LineChartSecondary;
            if (this.UsedChartTypes[iChartType])
            {
                C.LineChart lc = new C.LineChart();
                lc.Grouping = new C.Grouping() { Val = this.UsedChartOptions[iChartType].Grouping };
                lc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        lc.Append(this.DataSeries[i].ToLineChartSeries(IsStylish));
                    }
                }

                lc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (this.UsedChartOptions[iChartType].HasDropLines)
                {
                    lc.Append(this.UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));
                }

                lc.Append(new C.ShowMarker() { Val = this.UsedChartOptions[iChartType].ShowMarker });
                lc.Append(new C.Smooth() { Val = this.UsedChartOptions[iChartType].Smooth });

                lc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis1 });
                lc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis2 });

                pa.Append(lc);
            }
            #endregion

            #region Area3D
            iChartType = (int)SLDataSeriesChartType.Area3DChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.Area3DChart ac = new C.Area3DChart();
                ac.Grouping = new C.Grouping() { Val = this.UsedChartOptions[iChartType].Grouping };
                ac.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        ac.Append(this.DataSeries[i].ToAreaChartSeries(IsStylish));
                    }
                }

                ac.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (this.UsedChartOptions[iChartType].GapDepth != 150)
                {
                    ac.Append(new C.GapDepth() { Val = this.UsedChartOptions[iChartType].GapDepth });
                }

                ac.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                ac.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });
                ac.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis3 });

                pa.Append(ac);
            }
            #endregion

            #region Bar3D
            iChartType = (int)SLDataSeriesChartType.Bar3DChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.Bar3DChart bc = new C.Bar3DChart();
                bc.BarDirection = new C.BarDirection() { Val = this.UsedChartOptions[iChartType].BarDirection };
                bc.BarGrouping = new C.BarGrouping() { Val = this.UsedChartOptions[iChartType].BarGrouping };
                bc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        bc.Append(this.DataSeries[i].ToBarChartSeries(IsStylish));
                    }
                }

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                bc.Append(new C.GapWidth() { Val = this.UsedChartOptions[iChartType].GapWidth });

                if (this.UsedChartOptions[iChartType].GapDepth != 150)
                {
                    bc.Append(new C.GapDepth() { Val = this.UsedChartOptions[iChartType].GapDepth });
                }

                bc.Append(new C.Shape() { Val = this.UsedChartOptions[iChartType].Shape });

                bc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                bc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });
                bc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis3 });

                pa.Append(bc);
            }
            #endregion

            #region Bubble
            iChartType = (int)SLDataSeriesChartType.BubbleChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.BubbleChart bc = new C.BubbleChart();
                bc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        bc.Append(this.DataSeries[i].ToBubbleChartSeries(IsStylish));
                    }
                }

                bc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (!this.UsedChartOptions[iChartType].Bubble3D)
                {
                    bc.Append(new C.Bubble3D() { Val = this.UsedChartOptions[iChartType].Bubble3D });
                }

                if (this.UsedChartOptions[iChartType].BubbleScale != 100)
                {
                    bc.Append(new C.BubbleScale() { Val = this.UsedChartOptions[iChartType].BubbleScale });
                }

                if (!this.UsedChartOptions[iChartType].ShowNegativeBubbles)
                {
                    bc.Append(new C.ShowNegativeBubbles() { Val = this.UsedChartOptions[iChartType].ShowNegativeBubbles });
                }

                if (this.UsedChartOptions[iChartType].SizeRepresents != C.SizeRepresentsValues.Area)
                {
                    bc.Append(new C.SizeRepresents() { Val = this.UsedChartOptions[iChartType].SizeRepresents });
                }

                bc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                bc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });

                pa.Append(bc);
            }
            #endregion

            #region Line3D
            iChartType = (int)SLDataSeriesChartType.Line3DChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.Line3DChart lc = new C.Line3DChart();
                lc.Grouping = new C.Grouping() { Val = this.UsedChartOptions[iChartType].Grouping };
                lc.VaryColors = new C.VaryColors() { Val = false };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        lc.Append(this.DataSeries[i].ToLineChartSeries(IsStylish));
                    }
                }

                lc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (this.UsedChartOptions[iChartType].HasDropLines)
                {
                    lc.Append(this.UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));
                }

                if (this.UsedChartOptions[iChartType].GapDepth != 150)
                {
                    lc.Append(new C.GapDepth() { Val = this.UsedChartOptions[iChartType].GapDepth });
                }

                lc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                lc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });
                lc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis3 });

                pa.Append(lc);
            }
            #endregion

            #region Pie3D
            iChartType = (int)SLDataSeriesChartType.Pie3DChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.Pie3DChart pc = new C.Pie3DChart();
                pc.VaryColors = new C.VaryColors() { Val = this.UsedChartOptions[iChartType].VaryColors ?? true };

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        pc.Append(this.DataSeries[i].ToPieChartSeries(IsStylish));
                    }
                }

                pc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                pa.Append(pc);
            }
            #endregion

            #region Surface
            iChartType = (int)SLDataSeriesChartType.SurfaceChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.SurfaceChart sc = new C.SurfaceChart();
                if (this.UsedChartOptions[iChartType].bWireframe != null)
                {
                    sc.Wireframe = new C.Wireframe() { Val = this.UsedChartOptions[iChartType].Wireframe };
                }

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        sc.Append(this.DataSeries[i].ToSurfaceChartSeries(IsStylish));
                    }
                }

                sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });
                sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis3 });

                pa.Append(sc);
            }
            #endregion

            #region Surface3D
            iChartType = (int)SLDataSeriesChartType.Surface3DChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.Surface3DChart sc = new C.Surface3DChart();
                if (this.UsedChartOptions[iChartType].bWireframe != null)
                {
                    sc.Wireframe = new C.Wireframe() { Val = this.UsedChartOptions[iChartType].Wireframe };
                }

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        sc.Append(this.DataSeries[i].ToSurfaceChartSeries(IsStylish));
                    }
                }

                sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });
                sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis3 });

                pa.Append(sc);
            }
            #endregion

            #region Stock
            iChartType = (int)SLDataSeriesChartType.StockChart;
            if (this.UsedChartTypes[iChartType])
            {
                C.StockChart sc = new C.StockChart();

                for (i = 0; i < this.DataSeries.Count; ++i)
                {
                    if ((int)this.DataSeries[i].ChartType == iChartType)
                    {
                        sc.Append(this.DataSeries[i].ToLineChartSeries(IsStylish));
                    }
                }

                sc.Append(gdlo.ToDataLabels(new Dictionary<int, SLDataLabelOptions>(), false));

                if (this.UsedChartOptions[iChartType].HasDropLines)
                {
                    sc.Append(this.UsedChartOptions[iChartType].DropLines.ToDropLines(IsStylish));
                }

                if (this.UsedChartOptions[iChartType].HasHighLowLines)
                {
                    sc.Append(this.UsedChartOptions[iChartType].HighLowLines.ToHighLowLines(IsStylish));
                }

                if (this.UsedChartOptions[iChartType].HasUpDownBars)
                {
                    sc.Append(this.UsedChartOptions[iChartType].UpDownBars.ToUpDownBars(IsStylish));
                }

                // stock charts either have a bar chart as the primary chart (the Volume) or doesn't.
                // If there is, then it's either a Volume-High-Low-Close or Volumn-Open-High-Low-Close,
                // so we use the secondary axis IDs.
                if (this.UsedChartTypes[(int)SLDataSeriesChartType.BarChartColumnPrimary])
                {
                    sc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis1 });
                    sc.Append(new C.AxisId() { Val = SLConstants.SecondaryAxis2 });
                }
                else
                {
                    sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis1 });
                    sc.Append(new C.AxisId() { Val = SLConstants.PrimaryAxis2 });
                }

                pa.Append(sc);
            }
            #endregion

            if (this.HasPrimaryAxes)
            {
                this.PrimaryTextAxis.IsCrosses = this.PrimaryValueAxis.OtherAxisIsCrosses;
                this.PrimaryTextAxis.Crosses = this.PrimaryValueAxis.OtherAxisCrosses;
                this.PrimaryTextAxis.CrossesAt = this.PrimaryValueAxis.OtherAxisCrossesAt;

                this.PrimaryTextAxis.OtherAxisIsInReverseOrder = this.PrimaryValueAxis.InReverseOrder;

                if (this.PrimaryValueAxis.OtherAxisIsCrosses != null
                    && this.PrimaryValueAxis.OtherAxisIsCrosses.Value
                    && this.PrimaryValueAxis.OtherAxisCrosses == C.CrossesValues.Maximum)
                {
                    this.PrimaryTextAxis.OtherAxisCrossedAtMaximum = true;
                }
                else
                {
                    this.PrimaryTextAxis.OtherAxisCrossedAtMaximum = false;
                }

                this.PrimaryValueAxis.IsCrosses = this.PrimaryTextAxis.OtherAxisIsCrosses;
                this.PrimaryValueAxis.Crosses = this.PrimaryTextAxis.OtherAxisCrosses;
                this.PrimaryValueAxis.CrossesAt = this.PrimaryTextAxis.OtherAxisCrossesAt;

                this.PrimaryValueAxis.OtherAxisIsInReverseOrder = this.PrimaryTextAxis.InReverseOrder;

                if (this.PrimaryTextAxis.OtherAxisIsCrosses != null
                    && this.PrimaryTextAxis.OtherAxisIsCrosses.Value
                    && this.PrimaryTextAxis.OtherAxisCrosses == C.CrossesValues.Maximum)
                {
                    this.PrimaryValueAxis.OtherAxisCrossedAtMaximum = true;
                }
                else
                {
                    this.PrimaryValueAxis.OtherAxisCrossedAtMaximum = false;
                }

                switch (this.PrimaryTextAxis.AxisType)
                {
                    case SLAxisType.Category:
                        pa.Append(this.PrimaryTextAxis.ToCategoryAxis(IsStylish));
                        break;
                    case SLAxisType.Date:
                        pa.Append(this.PrimaryTextAxis.ToDateAxis(IsStylish));
                        break;
                    case SLAxisType.Value:
                        pa.Append(this.PrimaryTextAxis.ToValueAxis(IsStylish));
                        break;
                }
                pa.Append(this.PrimaryValueAxis.ToValueAxis(IsStylish));
            }

            if (this.HasDepthAxis)
            {
                pa.Append(this.DepthAxis.ToSeriesAxis(IsStylish));
            }

            if (this.HasSecondaryAxes)
            {
                this.SecondaryTextAxis.IsCrosses = this.SecondaryValueAxis.OtherAxisIsCrosses;
                this.SecondaryTextAxis.Crosses = this.SecondaryValueAxis.OtherAxisCrosses;
                this.SecondaryTextAxis.CrossesAt = this.SecondaryValueAxis.OtherAxisCrossesAt;

                this.SecondaryTextAxis.OtherAxisIsInReverseOrder = this.SecondaryValueAxis.InReverseOrder;

                if (this.SecondaryValueAxis.OtherAxisIsCrosses != null
                    && this.SecondaryValueAxis.OtherAxisIsCrosses.Value
                    && this.SecondaryValueAxis.OtherAxisCrosses == C.CrossesValues.Maximum)
                {
                    this.SecondaryTextAxis.OtherAxisCrossedAtMaximum = true;
                }
                else
                {
                    this.SecondaryTextAxis.OtherAxisCrossedAtMaximum = false;
                }

                this.SecondaryValueAxis.IsCrosses = this.SecondaryTextAxis.OtherAxisIsCrosses;
                this.SecondaryValueAxis.Crosses = this.SecondaryTextAxis.OtherAxisCrosses;
                this.SecondaryValueAxis.CrossesAt = this.SecondaryTextAxis.OtherAxisCrossesAt;

                this.SecondaryValueAxis.OtherAxisIsInReverseOrder = this.SecondaryTextAxis.InReverseOrder;

                if (this.SecondaryTextAxis.OtherAxisIsCrosses != null
                    && this.SecondaryTextAxis.OtherAxisIsCrosses.Value
                    && this.SecondaryTextAxis.OtherAxisCrosses == C.CrossesValues.Maximum)
                {
                    this.SecondaryValueAxis.OtherAxisCrossedAtMaximum = true;
                }
                else
                {
                    this.SecondaryValueAxis.OtherAxisCrossedAtMaximum = false;
                }

                // the order of axes is:
                // 1) primary category/date/value axis
                // 2) primary value axis
                // 3) secondary value axis
                // 4) secondary category/date/value axis
                pa.Append(this.SecondaryValueAxis.ToValueAxis(IsStylish));
                switch (this.SecondaryTextAxis.AxisType)
                {
                    case SLAxisType.Category:
                        pa.Append(this.SecondaryTextAxis.ToCategoryAxis(IsStylish));
                        break;
                    case SLAxisType.Date:
                        pa.Append(this.SecondaryTextAxis.ToDateAxis(IsStylish));
                        break;
                    case SLAxisType.Value:
                        pa.Append(this.SecondaryTextAxis.ToValueAxis(IsStylish));
                        break;
                }
            }

            if (this.ShowDataTable) pa.Append(this.DataTable.ToDataTable(IsStylish));

            if (this.ShapeProperties.HasShapeProperties) pa.Append(this.ShapeProperties.ToChartShapeProperties(IsStylish));

            return pa;
        }

        internal SLPlotArea Clone()
        {
            SLPlotArea pa = new SLPlotArea(this.ShapeProperties.listThemeColors, this.PrimaryTextAxis.Date1904, false, this.ShapeProperties.ThrowExceptionsIfAny);
            pa.InternalChartType = this.InternalChartType;

            int i;

            pa.UsedChartTypes = new bool[this.UsedChartTypes.Length];
            for (i = 0; i < this.UsedChartTypes.Length; ++i)
            {
                pa.UsedChartTypes[i] = this.UsedChartTypes[i];
            }

            pa.UsedChartOptions = new SLChartOptions[this.UsedChartOptions.Length];
            for (i = 0; i < this.UsedChartOptions.Length; ++i)
            {
                pa.UsedChartOptions[i] = this.UsedChartOptions[i].Clone();
            }

            pa.DataSeries = new List<SLDataSeries>();
            for (i = 0; i < this.DataSeries.Count; ++i)
            {
                pa.DataSeries.Add(this.DataSeries[i].Clone());
            }

            pa.Layout = this.Layout.Clone();
            pa.PrimaryTextAxis = this.PrimaryTextAxis.Clone();
            pa.PrimaryValueAxis = this.PrimaryValueAxis.Clone();
            pa.DepthAxis = this.DepthAxis.Clone();
            pa.SecondaryTextAxis = this.SecondaryTextAxis.Clone();
            pa.SecondaryValueAxis = this.SecondaryValueAxis.Clone();
            pa.HasPrimaryAxes = this.HasPrimaryAxes;
            pa.HasDepthAxis = this.HasDepthAxis;
            pa.HasSecondaryAxes = this.HasSecondaryAxes;
            pa.ShowDataTable = this.ShowDataTable;
            pa.DataTable = this.DataTable.Clone();
            pa.ShapeProperties = this.ShapeProperties.Clone();

            return pa;
        }
    }
}
