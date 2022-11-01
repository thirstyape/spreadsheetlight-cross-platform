using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    /// <summary>
    /// These correspond to the internal Open XML SDK classes
    /// </summary>
    internal enum SLInternalChartType
    {
        Area = 0,
        Area3D,
        Line,
        Line3D,
        Stock,
        Radar,
        Scatter,
        Pie,
        Pie3D,
        Doughnut,
        Bar,
        Bar3D,
        OfPie,
        Surface,
        Surface3D,
        Bubble
    }

    /// <summary>
    /// Data series display types.
    /// </summary>
    public enum SLChartDataDisplayType
    {
        /// <summary>
        /// Normal or clustered.
        /// </summary>
        Normal,
        /// <summary>
        /// Stacked.
        /// </summary>
        Stacked,
        /// <summary>
        /// 100% stacked.
        /// </summary>
        StackedMax
    }

    /// <summary>
    /// Built-in column chart types.
    /// </summary>
    public enum SLColumnChartType
    {
        /// <summary>
        /// Clustered Column.
        /// </summary>
        ClusteredColumn = 0,
        /// <summary>
        /// Stacked Column.
        /// </summary>
        StackedColumn,
        /// <summary>
        /// 100% Stacked Column.
        /// </summary>
        StackedColumnMax,
        /// <summary>
        /// 3D Clustered Column.
        /// </summary>
        ClusteredColumn3D,
        /// <summary>
        /// Stacked Column in 3D.
        /// </summary>
        StackedColumn3D,
        /// <summary>
        /// 100% Stacked Column in 3D.
        /// </summary>
        StackedColumnMax3D,
        /// <summary>
        /// 3D Column.
        /// </summary>
        Column3D,
        /// <summary>
        /// Clustered Cylinder.
        /// </summary>
        ClusteredCylinder,
        /// <summary>
        /// Stacked Cylinder.
        /// </summary>
        StackedCylinder,
        /// <summary>
        /// 100% Stacked Cylinder.
        /// </summary>
        StackedCylinderMax,
        /// <summary>
        /// 3D Cylinder.
        /// </summary>
        Cylinder3D,
        /// <summary>
        /// Clustered Cone.
        /// </summary>
        ClusteredCone,
        /// <summary>
        /// Stacked Cone.
        /// </summary>
        StackedCone,
        /// <summary>
        /// 100% Stacked Cone.
        /// </summary>
        StackedConeMax,
        /// <summary>
        /// 3D Cone.
        /// </summary>
        Cone3D,
        /// <summary>
        /// Clustered Pyramid.
        /// </summary>
        ClusteredPyramid,
        /// <summary>
        /// Stacked Pyramid.
        /// </summary>
        StackedPyramid,
        /// <summary>
        /// 100% Stacked Pyramid.
        /// </summary>
        StackedPyramidMax,
        /// <summary>
        /// 3D Pyramid.
        /// </summary>
        Pyramid3D
    }

    /// <summary>
    /// Built-in line chart types.
    /// </summary>
    public enum SLLineChartType
    {
        /// <summary>
        /// Line.
        /// </summary>
        Line = 0,
        /// <summary>
        /// Stacked Line.
        /// </summary>
        StackedLine,
        /// <summary>
        /// 100% Stacked Line.
        /// </summary>
        StackedLineMax,
        /// <summary>
        /// Line with Markers.
        /// </summary>
        LineWithMarkers,
        /// <summary>
        /// Stacked Line with Markers.
        /// </summary>
        StackedLineWithMarkers,
        /// <summary>
        /// 100% Stacked Line with Markers.
        /// </summary>
        StackedLineWithMarkersMax,
        /// <summary>
        /// 3D Line.
        /// </summary>
        Line3D
    }

    /// <summary>
    /// Built-in pie chart types.
    /// </summary>
    public enum SLPieChartType
    {
        /// <summary>
        /// Pie.
        /// </summary>
        Pie = 0,
        /// <summary>
        /// Pie in 3D.
        /// </summary>
        Pie3D,
        /// <summary>
        /// Pie of Pie.
        /// </summary>
        PieOfPie,
        /// <summary>
        /// Exploded Pie.
        /// </summary>
        ExplodedPie,
        /// <summary>
        /// Exploded Pie in 3D.
        /// </summary>
        ExplodedPie3D,
        /// <summary>
        /// Bar of Pie
        /// </summary>
        BarOfPie
    }

    /// <summary>
    /// Built-in bar chart types.
    /// </summary>
    public enum SLBarChartType
    {
        /// <summary>
        /// Clustered Bar.
        /// </summary>
        ClusteredBar = 0,
        /// <summary>
        /// Stacked Bar.
        /// </summary>
        StackedBar,
        /// <summary>
        /// 100% Stacked Bar.
        /// </summary>
        StackedBarMax,
        /// <summary>
        /// Clustered Bar in 3D.
        /// </summary>
        ClusteredBar3D,
        /// <summary>
        /// Stacked Bar in 3D.
        /// </summary>
        StackedBar3D,
        /// <summary>
        /// 100% Stacked Bar in 3D.
        /// </summary>
        StackedBarMax3D,
        /// <summary>
        /// Clustered Horizontal Cylinder.
        /// </summary>
        ClusteredHorizontalCylinder,
        /// <summary>
        /// Stacked Horizontal Cylinder.
        /// </summary>
        StackedHorizontalCylinder,
        /// <summary>
        /// 100% Stacked Horizontal Cylinder.
        /// </summary>
        StackedHorizontalCylinderMax,
        /// <summary>
        /// Clustered Horizontal Cone.
        /// </summary>
        ClusteredHorizontalCone,
        /// <summary>
        /// Stacked Horizontal Cone.
        /// </summary>
        StackedHorizontalCone,
        /// <summary>
        /// 100% Stacked Horizontal Cone.
        /// </summary>
        StackedHorizontalConeMax,
        /// <summary>
        /// Clustered Horizontal Pyramid.
        /// </summary>
        ClusteredHorizontalPyramid,
        /// <summary>
        /// Stacked Horizontal Pyramid.
        /// </summary>
        StackedHorizontalPyramid,
        /// <summary>
        /// 100% Stacked Horizontal Pyramid.
        /// </summary>
        StackedHorizontalPyramidMax
    }

    /// <summary>
    /// Built-in area chart types.
    /// </summary>
    public enum SLAreaChartType
    {
        /// <summary>
        /// Area.
        /// </summary>
        Area = 0,
        /// <summary>
        /// Stacked Area.
        /// </summary>
        StackedArea,
        /// <summary>
        /// 100% Stacked Area.
        /// </summary>
        StackedAreaMax,
        /// <summary>
        /// 3D Area.
        /// </summary>
        Area3D,
        /// <summary>
        /// Stacked Area in 3D.
        /// </summary>
        StackedArea3D,
        /// <summary>
        /// 100% Stacked Area in 3D.
        /// </summary>
        StackedAreaMax3D
    }

    /// <summary>
    /// Built-in scatter chart types.
    /// </summary>
    public enum SLScatterChartType
    {
        /// <summary>
        /// Scatter with only Markers.
        /// </summary>
        ScatterWithOnlyMarkers = 0,
        /// <summary>
        /// Scatter with Smooth Lines and Markers.
        /// </summary>
        ScatterWithSmoothLinesAndMarkers,
        /// <summary>
        /// Scatter with Smooth Lines.
        /// </summary>
        ScatterWithSmoothLines,
        /// <summary>
        /// Scatter with Straight Lines and Markers.
        /// </summary>
        ScatterWithStraightLinesAndMarkers,
        /// <summary>
        /// Scatter with Straight Lines.
        /// </summary>
        ScatterWithStraightLines
    }

    /// <summary>
    /// Built-in stock chart types.
    /// </summary>
    public enum SLStockChartType
    {
        /// <summary>
        /// High-Low-Close.
        /// </summary>
        HighLowClose = 0,
        /// <summary>
        /// Open-High-Low-Close.
        /// </summary>
        OpenHighLowClose,
        /// <summary>
        /// Volume-High-Low-Close.
        /// </summary>
        VolumeHighLowClose,
        /// <summary>
        /// Volume-Open-High-Low-Close.
        /// </summary>
        VolumeOpenHighLowClose
    }

    /// <summary>
    /// Built-in surface chart types.
    /// </summary>
    public enum SLSurfaceChartType
    {
        /// <summary>
        /// 3D Surface.
        /// </summary>
        Surface3D = 0,
        /// <summary>
        /// Wiredframe 3D Surface.
        /// </summary>
        WireframeSurface3D,
        /// <summary>
        /// Contour.
        /// </summary>
        Contour,
        /// <summary>
        /// Wireframe Contour.
        /// </summary>
        WireframeContour
    }

    /// <summary>
    /// Built-in doughnut chart types.
    /// </summary>
    public enum SLDoughnutChartType
    {
        /// <summary>
        /// Doughnut.
        /// </summary>
        Doughnut = 0,
        /// <summary>
        /// Exploded Doughnut.
        /// </summary>
        ExplodedDoughnut
    }

    /// <summary>
    /// Built-in bubble chart types.
    /// </summary>
    public enum SLBubbleChartType
    {
        /// <summary>
        /// Bubble.
        /// </summary>
        Bubble = 0,
        /// <summary>
        /// Bubble with a 3D effect.
        /// </summary>
        Bubble3D
    }

    /// <summary>
    /// Built-in radar chart types.
    /// </summary>
    public enum SLRadarChartType
    {
        /// <summary>
        /// Radar.
        /// </summary>
        Radar = 0,
        /// <summary>
        /// Radar with Markers.
        /// </summary>
        RadarWithMarkers,
        /// <summary>
        /// Filled Radar.
        /// </summary>
        FilledRadar
    }

    /// <summary>
    /// Built-in chart styles.
    /// </summary>
    public enum SLChartStyle : byte
    {
        // the numbers assigned have to be those assigned as follows.

        /// <summary>
        /// Standard style in black and white.
        /// </summary>
        Style1 = 1,
        /// <summary>
        /// Standard style in theme colors. This is the default.
        /// </summary>
        Style2 = 2,
        /// <summary>
        /// Standard style in tints of accent 1 color.
        /// </summary>
        Style3 = 3,
        /// <summary>
        /// Standard style in tints of accent 2 color.
        /// </summary>
        Style4 = 4,
        /// <summary>
        /// Standard style in tints of accent 3 color.
        /// </summary>
        Style5 = 5,
        /// <summary>
        /// Standard style in tints of accent 4 color.
        /// </summary>
        Style6 = 6,
        /// <summary>
        /// Standard style in tints of accent 5 color.
        /// </summary>
        Style7 = 7,
        /// <summary>
        /// Standard style in tints of accent 6 color.
        /// </summary>
        Style8 = 8,
        /// <summary>
        /// Bordered data series in black and white.
        /// </summary>
        Style9 = 9,
        /// <summary>
        /// Bordered data series in theme colors.
        /// </summary>
        Style10 = 10,
        /// <summary>
        /// Bordered data series in tints of accent 1 color.
        /// </summary>
        Style11 = 11,
        /// <summary>
        /// Bordered data series in tints of accent 2 color.
        /// </summary>
        Style12 = 12,
        /// <summary>
        /// Bordered data series in tints of accent 3 color.
        /// </summary>
        Style13 = 13,
        /// <summary>
        /// Bordered data series in tints of accent 4 color.
        /// </summary>
        Style14 = 14,
        /// <summary>
        /// Bordered data series in tints of accent 5 color.
        /// </summary>
        Style15 = 15,
        /// <summary>
        /// Bordered data series in tints of accent 6 color.
        /// </summary>
        Style16 = 16,
        /// <summary>
        /// Softly blurred data series in black and white.
        /// </summary>
        Style17 = 17,
        /// <summary>
        /// Softly blurred data series in theme colors.
        /// </summary>
        Style18 = 18,
        /// <summary>
        /// Softly blurred data series in tints of accent 1 color.
        /// </summary>
        Style19 = 19,
        /// <summary>
        /// Softly blurred data series in tints of accent 2 color.
        /// </summary>
        Style20 = 20,
        /// <summary>
        /// Softly blurred data series in tints of accent 3 color.
        /// </summary>
        Style21 = 21,
        /// <summary>
        /// Softly blurred data series in tints of accent 4 color.
        /// </summary>
        Style22 = 22,
        /// <summary>
        /// Softly blurred data series in tints of accent 5 color.
        /// </summary>
        Style23 = 23,
        /// <summary>
        /// Softly blurred data series in tints of accent 6 color.
        /// </summary>
        Style24 = 24,
        /// <summary>
        /// Bevelled data series in black and white.
        /// </summary>
        Style25 = 25,
        /// <summary>
        /// Bevelled data series in theme colors.
        /// </summary>
        Style26 = 26,
        /// <summary>
        /// Bevelled data series in tints of accent 1 color.
        /// </summary>
        Style27 = 27,
        /// <summary>
        /// Bevelled data series in tints of accent 2 color.
        /// </summary>
        Style28 = 28,
        /// <summary>
        /// Bevelled data series in tints of accent 3 color.
        /// </summary>
        Style29 = 29,
        /// <summary>
        /// Bevelled data series in tints of accent 4 color.
        /// </summary>
        Style30 = 30,
        /// <summary>
        /// Bevelled data series in tints of accent 5 color.
        /// </summary>
        Style31 = 31,
        /// <summary>
        /// Bevelled data series in tints of accent 6 color.
        /// </summary>
        Style32 = 32,
        /// <summary>
        /// Standard style in black and white, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style33 = 33,
        /// <summary>
        /// Standard style in theme colors, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style34 = 34,
        /// <summary>
        /// Standard style in tints of accent 1 color, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style35 = 35,
        /// <summary>
        /// Standard style in tints of accent 2 color, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style36 = 36,
        /// <summary>
        /// Standard style in tints of accent 3 color, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style37 = 37,
        /// <summary>
        /// Standard style in tints of accent 4 color, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style38 = 38,
        /// <summary>
        /// Standard style in tints of accent 5 color, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style39 = 39,
        /// <summary>
        /// Standard style in tints of accent 6 color, with gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style40 = 40,
        /// <summary>
        /// Softly blurred and bevelled data series in black and white, with black chart area and gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style41 = 41,
        /// <summary>
        /// Softly blurred and bevelled data series in theme colors, with black chart area and gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style42 = 42,
        /// <summary>
        /// Softly blurred and bevelled data series in tints of accent 1 color, with black chart area and gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style43 = 43,
        /// <summary>
        /// Softly blurred and bevelled data series in tints of accent 2 color, with black chart area and gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style44 = 44,
        /// <summary>
        /// Softly blurred and bevelled data series in tints of accent 3 color, with black chart area and gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style45 = 45,
        /// <summary>
        /// Softly blurred and bevelled data series in tints of accent 4 color, with black chart area and gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style46 = 46,
        /// <summary>
        /// Softly blurred and bevelled data series in tints of accent 5 color, with black chart area and gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style47 = 47,
        /// <summary>
        /// Softly blurred and bevelled data series in tints of accent 6 color, with black chart area and gray-filled plot area (side wall, back wall and floor for 3D charts).
        /// </summary>
        Style48 = 48
    }

    // this is for ChartSpace root class

    /// <summary>
    /// Encapsulates properties and methods for a chart to be inserted into a worksheet.
    /// </summary>
    public class SLChart
    {
        internal List<System.Drawing.Color> listThemeColors;
        internal bool ThrowExceptionsIfAny { get; set; }
        internal bool Date1904 { get; set; }

        /// <summary>
        /// True if follow latest Excel styling defaults (but no guarantees because I might not
        /// be able to afford to keep buying latest Office/Excel).
        /// </summary>
        internal bool IsStylish { get; set; }

        /// <summary>
        /// Specifies whether the chart has rounded corners. In Microsoft Excel, you might find this setting under "Border Styles" when formatting the chart area.
        /// </summary>
        public bool RoundedCorners { get; set; }

        internal bool IsCombinable { get; set; }

        internal double TopPosition { get; set; }
        internal double LeftPosition { get; set; }
        internal double BottomPosition { get; set; }
        internal double RightPosition { get; set; }

        // this is the primary data source
        internal string WorksheetName { get; set; }
        internal bool RowsAsDataSeries { get; set; }
        internal bool ShowHiddenData { get; set; }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal SLChartStyle ChartStyle { get; set; }

        /// <summary>
        /// The default is to show empty cells with a gap (or whichever option is appropriate for the chart). Note that "Zero" and "Span" are used mostly for line, scatter and radar charts. Use "Zero" to force a zero value, and "Span" to connect data points across the empty cell.
        /// </summary>
        public C.DisplayBlanksAsValues ShowEmptyCellsAs { get; set; }

        /// <summary>
        /// Indicates whether data labels over the maximum value of the chart is shown. The default value is true.
        /// </summary>
        public bool ShowDataLabelsOverMaximum { get; set; }

        internal bool HasView3D
        {
            get
            {
                return this.RotateX != null || this.HeightPercent != null || this.RotateY != null || this.DepthPercent != null || this.RightAngleAxes != null || this.Perspective != null;
            }
        }
        // RotateX and RotateY don't correspond to the X- and Y-axis on the Excel user interface.
        // Why? Why?!? WHY?!?! I don't know. Go ask Microsoft...
        internal sbyte? RotateX { get; set; }
        internal ushort? HeightPercent { get; set; }
        internal ushort? RotateY { get; set; }
        internal ushort? DepthPercent { get; set; }
        internal bool? RightAngleAxes { get; set; }
        /// <summary>
        /// This is double that's shown in Excel. Excel values range from 0 to 120 degrees.
        /// So this is 0 to 240 units. "Default" rotation angle is 30 (15 degrees).
        /// Did Microsoft want to make full use of the byte range value?
        /// </summary>
        internal byte? Perspective { get; set; }

        /// <summary>
        /// A friendly name for the chart. By default, this is in the form of "Chart #", where "#" is a number.
        /// </summary>
        public string ChartName { get; set; }

        internal bool HasTitle { get; set; }
        /// <summary>
        /// The chart title. By default the chart title is hidden, so make sure to show it if chart title properties are set.
        /// </summary>
        public SLTitle Title { get; set; }

        internal bool Is3D;

        /// <summary>
        /// The floor of 3D charts.
        /// </summary>
        public SLFloor Floor { get; set; }

        /// <summary>
        /// The side wall of 3D charts. Note that contour charts don't show the side wall, even though they're technically 3D charts.
        /// </summary>
        public SLSideWall SideWall { get; set; }

        /// <summary>
        /// The back wall of 3D charts. Note that contour charts don't show the back wall, even though they're technically 3D charts.
        /// </summary>
        public SLBackWall BackWall { get; set; }

        /// <summary>
        /// The plot area.
        /// </summary>
        public SLPlotArea PlotArea { get; set; }

        /// <summary>
        /// The primary chart text axis. This is usually the horizontal axis at the bottom (bar charts have them on the left).
        /// Depending on the type of chart, this can be a category, date or value axis.
        /// </summary>
        public SLTextAxis PrimaryTextAxis { get { return this.PlotArea.PrimaryTextAxis; } }

        /// <summary>
        /// The primary chart value axis. This is usually the vertical axis on the left (bar charts have them at the bottom).
        /// </summary>
        public SLValueAxis PrimaryValueAxis { get { return this.PlotArea.PrimaryValueAxis; } }

        /// <summary>
        /// The depth axis for 3D charts.
        /// </summary>
        public SLSeriesAxis DepthAxis { get { return this.PlotArea.DepthAxis; } }

        /// <summary>
        /// The secondary chart text axis. This is usually the horizontal axis at the top (bar charts have them on the left initially until you show this axis).
        /// Depending on the type of chart, this can be a category, date or value axis.
        /// </summary>
        public SLTextAxis SecondaryTextAxis { get { return this.PlotArea.SecondaryTextAxis; } }

        /// <summary>
        /// The secondary chart value axis. This is usually the vertical axis on the right (bar charts have them at the top).
        /// </summary>
        public SLValueAxis SecondaryValueAxis { get { return this.PlotArea.SecondaryValueAxis; } }

        // used to moderate the behaviour of the secondary text axis.
        // Initially, the axis is at the bottom (if not bar chart), but deleted.
        // Then when shown, the axis goes to the top. If then hidden, the axis stays at the top but is deleted.
        internal bool HasShownSecondaryTextAxis;

        /// <summary>
        /// Specifies if the data table is shown.
        /// </summary>
        public bool ShowDataTable
        {
            get { return this.PlotArea.ShowDataTable; }
            set { this.PlotArea.ShowDataTable = value; }
        }

        /// <summary>
        /// The data table of the chart.
        /// </summary>
        public SLDataTable DataTable { get { return this.PlotArea.DataTable; } }

        internal bool ShowLegend { get; set; }
        /// <summary>
        /// The chart legend.
        /// </summary>
        public SLLegend Legend { get; set; }

        internal SLA.SLShapeProperties ShapeProperties;

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

        internal SLChart()
        {
        }

        /// <summary>
        /// Set the chart style using one of the built-in styles. WARNING: This is supposedly phased out in Excel 2013. Maybe it'll be replaced by something else, maybe not at all.
        /// </summary>
        /// <param name="ChartStyle">A built-in chart style.</param>
        public void SetChartStyle(SLChartStyle ChartStyle)
        {
            this.ChartStyle = ChartStyle;
        }

        /// <summary>
        /// Set a pie chart using one of the built-in pie chart types.
        /// </summary>
        /// <param name="ChartType">A built-in pie chart type.</param>
        public void SetChartType(SLPieChartType ChartType)
        {
            this.SetChartType(ChartType, null);
        }

        /// <summary>
        /// Set a pie chart using one of the built-in pie chart types.
        /// </summary>
        /// <param name="ChartType">A built-in pie chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLPieChartType ChartType, SLPieChartOptions Options)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLPieChartType.Pie:
                    vType = SLDataSeriesChartType.PieChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    this.PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);
                    break;
                case SLPieChartType.Pie3D:
                    this.RotateX = 30;
                    if (Options != null)
                    {
                        this.RotateY = Options.FirstSliceAngle;
                    }
                    else
                    {
                        this.RotateY = 0;
                    }
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Pie3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);
                    break;
                case SLPieChartType.PieOfPie:
                    vType = SLDataSeriesChartType.OfPieChartPie;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    this.PlotArea.UsedChartOptions[iChartType].GapWidth = 100;
                    this.PlotArea.UsedChartOptions[iChartType].SecondPieSize = 75;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);
                    break;
                case SLPieChartType.ExplodedPie:
                    vType = SLDataSeriesChartType.PieChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    this.PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Explosion = 25;
                    }
                    break;
                case SLPieChartType.ExplodedPie3D:
                    this.RotateX = 30;
                    if (Options != null)
                    {
                        this.RotateY = Options.FirstSliceAngle;
                    }
                    else
                    {
                        this.RotateY = 0;
                    }
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Pie3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Explosion = 25;
                    }
                    break;
                case SLPieChartType.BarOfPie:
                    vType = SLDataSeriesChartType.OfPieChartBar;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    this.PlotArea.UsedChartOptions[iChartType].GapWidth = 100;
                    this.PlotArea.UsedChartOptions[iChartType].SecondPieSize = 75;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);
                    break;
            }
        }

        /// <summary>
        /// Set a surface chart using one of the built-in surface chart types.
        /// </summary>
        /// <param name="ChartType">A built-in surface chart type.</param>
        public void SetChartType(SLSurfaceChartType ChartType)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLSurfaceChartType.Surface3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Surface3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    this.PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
                    this.PlotArea.HasDepthAxis = true;
                    this.PlotArea.DepthAxis.IsCrosses = true;
                    this.PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
                case SLSurfaceChartType.WireframeSurface3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Surface3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Wireframe = true;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    this.PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
                    this.PlotArea.HasDepthAxis = true;
                    this.PlotArea.DepthAxis.IsCrosses = true;
                    this.PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
                case SLSurfaceChartType.Contour:
                    this.RotateX = 90;
                    this.RotateY = 0;
                    this.RightAngleAxes = false;
                    this.Perspective = 0;

                    vType = SLDataSeriesChartType.SurfaceChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    this.PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
                    this.PlotArea.HasDepthAxis = true;
                    this.PlotArea.DepthAxis.IsCrosses = true;
                    this.PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
                case SLSurfaceChartType.WireframeContour:
                    this.RotateX = 90;
                    this.RotateY = 0;
                    this.RightAngleAxes = false;
                    this.Perspective = 0;

                    vType = SLDataSeriesChartType.Surface3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Wireframe = true;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    this.PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
                    this.PlotArea.HasDepthAxis = true;
                    this.PlotArea.DepthAxis.IsCrosses = true;
                    this.PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
            }
        }

        /// <summary>
        /// Set a radar chart using one of the built-in radar chart types.
        /// </summary>
        /// <param name="ChartType">A built-in radar chart type.</param>
        public void SetChartType(SLRadarChartType ChartType)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLRadarChartType.Radar:
                    vType = SLDataSeriesChartType.RadarChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Marker;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;
                    }

                    if (this.IsStylish)
                    {
                        this.Legend.LegendPosition = C.LegendPositionValues.Top;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    this.PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;
                    break;
                case SLRadarChartType.RadarWithMarkers:
                    vType = SLDataSeriesChartType.RadarChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Marker;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    if (this.IsStylish)
                    {
                        for (int i = 0; i < this.PlotArea.DataSeries.Count; ++i)
                        {
                            this.PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            this.PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }
                        this.Legend.LegendPosition = C.LegendPositionValues.Top;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    this.PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;
                    break;
                case SLRadarChartType.FilledRadar:
                    vType = SLDataSeriesChartType.RadarChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Filled;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    if (this.IsStylish)
                    {
                        this.Legend.LegendPosition = C.LegendPositionValues.Top;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    this.PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;
                    break;
            }
        }

        // TODO! What's the correct bubble chart behaviour?

        /// <summary>
        /// Set a bubble chart using one of the built-in bubble chart types.
        /// </summary>
        /// <param name="ChartType">A built-in bubble chart type.</param>
        public void SetChartType(SLBubbleChartType ChartType)
        {
            this.SetChartType(ChartType, null);
        }

        /// <summary>
        /// Set a bubble chart using one of the built-in bubble chart types.
        /// </summary>
        /// <param name="ChartType">A built-in bubble chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLBubbleChartType ChartType, SLBubbleChartOptions Options)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            int i, index;
            SLNumberLiteral nl = new SLNumberLiteral();
            nl.FormatCode = SLConstants.NumberFormatGeneral;
            nl.PointCount = (uint)(this.EndRowIndex - this.StartRowIndex);
            for (i = 0; i < (this.EndRowIndex - this.StartRowIndex); ++i)
            {
                nl.Points.Add(new SLNumericPoint() { Index = (uint)i, NumericValue = "1" });
            }

            double fTemp = 0;

            List<SLDataSeries> series = new List<SLDataSeries>();
            SLDataSeries ser;
            for (index = 0, i = 0; i < this.PlotArea.DataSeries.Count; ++index, ++i)
            {
                ser = new SLDataSeries(this.listThemeColors, this.ThrowExceptionsIfAny);
                ser.Index = (uint)index;
                ser.Order = (uint)index;
                ser.IsStringReference = null;
                ser.StringReference = this.PlotArea.DataSeries[i].StringReference.Clone();

                ser.AxisData = this.PlotArea.DataSeries[i].AxisData.Clone();
                if (this.PlotArea.DataSeries[i].StringReference.Points.Count > 0)
                {
                    foreach (var pt in ser.AxisData.StringReference.Points)
                    {
                        ++pt.Index;
                    }

                    // move one row up
                    --ser.AxisData.StringReference.StartRowIndex;
                    ser.AxisData.StringReference.RefreshFormula();

                    ser.AxisData.StringReference.Points.Insert(0,
                        this.PlotArea.DataSeries[i].StringReference.Points[0].Clone());
                    ++ser.AxisData.StringReference.PointCount;
                }
                ser.NumberData = this.PlotArea.DataSeries[i].NumberData.Clone();
                if (this.PlotArea.DataSeries[i].StringReference.Points.Count > 1)
                {
                    foreach (var pt in ser.NumberData.NumberReference.NumberingCache.Points)
                    {
                        ++pt.Index;
                    }

                    --ser.NumberData.NumberReference.StartRowIndex;
                    ser.NumberData.NumberReference.RefreshFormula();

                    if (double.TryParse(this.PlotArea.DataSeries[i].StringReference.Points[1].NumericValue, NumberStyles.Any, CultureInfo.InvariantCulture, out fTemp))
                    {
                        ser.NumberData.NumberReference.NumberingCache.Points.Insert(0, new SLNumericPoint() { Index = 0, NumericValue = fTemp.ToString(CultureInfo.InvariantCulture) });
                    }
                    else
                    {
                        ser.NumberData.NumberReference.NumberingCache.Points.Insert(0, new SLNumericPoint() { Index = 0, NumericValue = "0" });
                    }
                    ++ser.NumberData.NumberReference.NumberingCache.PointCount;
                }

                ++i;
                if (i < this.PlotArea.DataSeries.Count)
                {
                    ser.BubbleSize = this.PlotArea.DataSeries[i].NumberData.Clone();

                    if (this.PlotArea.DataSeries[i].StringReference.Points.Count > 2)
                    {
                        foreach (var pt in ser.BubbleSize.NumberReference.NumberingCache.Points)
                        {
                            ++pt.Index;
                        }

                        --ser.BubbleSize.NumberReference.StartRowIndex;
                        ser.BubbleSize.NumberReference.RefreshFormula();

                        if (double.TryParse(this.PlotArea.DataSeries[i].StringReference.Points[2].NumericValue, NumberStyles.Any, CultureInfo.InvariantCulture, out fTemp))
                        {
                            ser.BubbleSize.NumberReference.NumberingCache.Points.Insert(0, new SLNumericPoint() { Index = 0, NumericValue = fTemp.ToString(CultureInfo.InvariantCulture) });
                        }
                        else
                        {
                            ser.BubbleSize.NumberReference.NumberingCache.Points.Insert(0, new SLNumericPoint() { Index = 0, NumericValue = "0" });
                        }
                        ++ser.BubbleSize.NumberReference.NumberingCache.PointCount;
                    }
                }
                else
                {
                    ser.BubbleSize.UseNumberLiteral = true;
                    ser.BubbleSize.NumberLiteral = nl.Clone();
                }
                series.Add(ser);
            }

            this.PlotArea.DataSeries = series;

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLBubbleChartType.Bubble:
                    vType = SLDataSeriesChartType.BubbleChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BubbleScale = 100;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Bubble3D = false;
                    }

                    this.SetPlotAreaValueAxes();
                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLBubbleChartType.Bubble3D:
                    vType = SLDataSeriesChartType.BubbleChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BubbleScale = 100;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Bubble3D = true;
                    }

                    this.SetPlotAreaValueAxes();
                    this.PlotArea.HasPrimaryAxes = true;
                    break;
            }
        }

        /// <summary>
        /// Set a stock chart using one of the built-in stock chart types.
        /// </summary>
        /// <param name="ChartType">A built-in stock chart type.</param>
        public void SetChartType(SLStockChartType ChartType)
        {
            this.SetChartType(ChartType, null);
        }

        /// <summary>
        /// Set a stock chart using one of the built-in stock chart types.
        /// </summary>
        /// <param name="ChartType">A built-in stock chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLStockChartType ChartType, SLStockChartOptions Options)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            int i;
            int iBarChartType;

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLStockChartType.HighLowClose:
                    vType = SLDataSeriesChartType.StockChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                    for (i = 0; i < this.PlotArea.DataSeries.Count; ++i)
                    {
                        this.PlotArea.DataSeries[i].ChartType = vType;
                        if (this.IsStylish)
                        {
                            this.PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                            this.PlotArea.DataSeries[i].Options.Line.CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Round;
                            this.PlotArea.DataSeries[i].Options.Line.SetNoLine();
                            this.PlotArea.DataSeries[i].Options.Line.JoinType = SLA.SLLineJoinValues.Round;
                            this.PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        }
                    }

                    // this is for Close
                    if (this.PlotArea.DataSeries.Count > 2)
                    {
                        this.PlotArea.DataSeries[2].Options.Marker.Symbol = C.MarkerStyleValues.Dot;
                        this.PlotArea.DataSeries[2].Options.Marker.Size = 3;
                        if (IsStylish)
                        {
                            this.PlotArea.DataSeries[2].Options.Marker.Fill.SetSolidFill(A.SchemeColorValues.Accent3, 0, 0);
                            this.PlotArea.DataSeries[2].Options.Marker.Line.Width = 0.75m;
                            this.PlotArea.DataSeries[2].Options.Marker.Line.SetSolidLine(A.SchemeColorValues.Accent3, 0, 0);
                        }
                    }

                    this.PlotArea.UsedChartOptions[iChartType].HasHighLowLines = true;
                    if (this.IsStylish)
                    {
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Width = 0.75m;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CapType = A.LineCapValues.Flat;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.25m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.JoinType = SLA.SLLineJoinValues.Round;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = false;
                        this.PlotArea.PrimaryTextAxis.Fill.SetNoFill();
                        // 2.25 pt width
                        this.PlotArea.PrimaryTextAxis.Line.Width = 0.75m;
                        this.PlotArea.PrimaryTextAxis.Line.CapType = A.LineCapValues.Flat;
                        this.PlotArea.PrimaryTextAxis.Line.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.PrimaryTextAxis.Line.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.PrimaryTextAxis.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                        this.PlotArea.PrimaryTextAxis.Line.JoinType = SLA.SLLineJoinValues.Round;
                        SLRstType rst = new SLRstType();
                        rst.AppendText(" ", new SLFont()
                        {
                            FontScheme = DocumentFormat.OpenXml.Spreadsheet.FontSchemeValues.Minor,
                            FontSize = 9,
                            Bold = false,
                            Italic = false,
                            Underline = DocumentFormat.OpenXml.Spreadsheet.UnderlineValues.None,
                            Strike = false
                        });
                        this.PlotArea.PrimaryTextAxis.Title.SetTitle(rst);
                        this.PlotArea.PrimaryTextAxis.Title.Fill.SetSolidFill(A.SchemeColorValues.Text1, 0.35m, 0);

                        this.PlotArea.PrimaryValueAxis.MinorTickMark = C.TickMarkValues.None;
                    }

                    if (this.IsStylish) this.Legend.LegendPosition = C.LegendPositionValues.Bottom;
                    this.PlotArea.SetDataSeriesAutoAxisType();
                    this.ShowEmptyCellsAs = C.DisplayBlanksAsValues.Gap;
                    break;
                case SLStockChartType.OpenHighLowClose:
                    vType = SLDataSeriesChartType.StockChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                    for (i = 0; i < this.PlotArea.DataSeries.Count; ++i)
                    {
                        this.PlotArea.DataSeries[i].ChartType = vType;
                        if (this.IsStylish)
                        {
                            // 2.25 pt width
                            this.PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                            this.PlotArea.DataSeries[i].Options.Line.CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Round;
                            this.PlotArea.DataSeries[i].Options.Line.SetNoLine();
                            this.PlotArea.DataSeries[i].Options.Line.JoinType = SLA.SLLineJoinValues.Round;
                            this.PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        }
                    }

                    this.PlotArea.UsedChartOptions[iChartType].HasHighLowLines = true;
                    this.PlotArea.UsedChartOptions[iChartType].HasUpDownBars = true;
                    if (this.IsStylish)
                    {
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Width = 0.75m;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CapType = A.LineCapValues.Flat;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.25m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.JoinType = SLA.SLLineJoinValues.Round;

                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.GapWidth = 150;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Fill.SetSolidFill(A.SchemeColorValues.Light1, 0, 0);
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.Width = 0.75m;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.CapType = A.LineCapValues.Flat;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.SetSolidLine(A.SchemeColorValues.Text1, 0.35m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.JoinType = SLA.SLLineJoinValues.Round;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Fill.SetSolidFill(A.SchemeColorValues.Dark1, 0.25m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.Width = 0.75m;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.CapType = A.LineCapValues.Flat;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.SetSolidLine(A.SchemeColorValues.Text1, 0.35m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.JoinType = SLA.SLLineJoinValues.Round;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;

                    if (this.IsStylish) this.Legend.LegendPosition = C.LegendPositionValues.Bottom;
                    this.PlotArea.SetDataSeriesAutoAxisType();
                    this.ShowEmptyCellsAs = C.DisplayBlanksAsValues.Gap;
                    break;
                case SLStockChartType.VolumeHighLowClose:
                    vType = SLDataSeriesChartType.StockChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                    iBarChartType = (int)SLDataSeriesChartType.BarChartColumnPrimary;
                    for (i = 0; i < this.PlotArea.DataSeries.Count; ++i)
                    {
                        if (i == 0)
                        {
                            this.PlotArea.DataSeries[i].ChartType = SLDataSeriesChartType.BarChartColumnPrimary;
                            if (this.IsStylish)
                            {
                                this.PlotArea.DataSeries[i].Options.Fill.SetSolidFill(A.SchemeColorValues.Accent1, 0, 0);
                                // 2.25 pt width
                                this.PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                                this.PlotArea.DataSeries[i].Options.Line.SetNoLine();
                            }

                            this.PlotArea.UsedChartTypes[iBarChartType] = true;
                            this.PlotArea.UsedChartOptions[iBarChartType].BarDirection = C.BarDirectionValues.Column;
                            this.PlotArea.UsedChartOptions[iBarChartType].BarGrouping = C.BarGroupingValues.Clustered;
                            if (Options != null)
                            {
                                this.PlotArea.UsedChartOptions[iBarChartType].GapWidth = Options.GapWidth;
                                this.PlotArea.UsedChartOptions[iBarChartType].Overlap = Options.Overlap;
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.Fill = Options.Fill.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.Outline = Options.Border.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.Shadow = Options.Shadow.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.Glow = Options.Glow.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.SoftEdge = Options.SoftEdge.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.Format3D = Options.Format3D.Clone();
                            }
                        }
                        else
                        {
                            this.PlotArea.DataSeries[i].ChartType = vType;
                            if (this.IsStylish)
                            {
                                // 2.25 pt width
                                this.PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                                this.PlotArea.DataSeries[i].Options.Line.CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Round;
                                this.PlotArea.DataSeries[i].Options.Line.SetNoLine();
                                this.PlotArea.DataSeries[i].Options.Line.JoinType = SLA.SLLineJoinValues.Round;
                                this.PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.None;
                            }
                        }
                    }

                    // this is for Close
                    if (this.PlotArea.DataSeries.Count > 3)
                    {
                        this.PlotArea.DataSeries[3].Options.Marker.Symbol = C.MarkerStyleValues.Dot;
                        this.PlotArea.DataSeries[3].Options.Marker.Size = 5;
                        if (IsStylish)
                        {
                            this.PlotArea.DataSeries[3].Options.Marker.Fill.SetSolidFill(A.SchemeColorValues.Accent4, 0, 0);
                            this.PlotArea.DataSeries[3].Options.Marker.Line.Width = 0.75m;
                            this.PlotArea.DataSeries[3].Options.Marker.Line.SetSolidLine(A.SchemeColorValues.Accent4, 0, 0);
                        }
                    }

                    this.PlotArea.UsedChartOptions[iChartType].HasHighLowLines = true;
                    if (this.IsStylish)
                    {
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Width = 0.75m;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CapType = A.LineCapValues.Flat;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.25m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.JoinType = SLA.SLLineJoinValues.Round;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
                    this.PlotArea.HasSecondaryAxes = true;
                    this.PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;
                    this.PlotArea.SecondaryValueAxis.ForceAxisPosition = true;
                    this.PlotArea.SecondaryValueAxis.IsCrosses = true;
                    this.PlotArea.SecondaryValueAxis.Crosses = C.CrossesValues.Maximum;
                    this.PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
                    //this.PlotArea.SecondaryValueAxis.OtherAxisIsCrosses = true;
                    //this.PlotArea.SecondaryValueAxis.OtherAxisCrosses = C.CrossesValues.AutoZero;
                    this.PlotArea.SecondaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    this.PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
                    //this.PlotArea.SecondaryTextAxis.IsCrosses = true;
                    //this.PlotArea.SecondaryTextAxis.Crosses = C.CrossesValues.AutoZero;
                    this.PlotArea.SecondaryTextAxis.OtherAxisIsCrosses = true;
                    this.PlotArea.SecondaryTextAxis.OtherAxisCrosses = C.CrossesValues.Maximum;

                    if (this.IsStylish)
                    {
                        this.PlotArea.SecondaryValueAxis.ShowMajorGridlines = true;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.Width = 0.75m;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.CapType = A.LineCapValues.Flat;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.JoinType = SLA.SLLineJoinValues.Round;

                        this.PlotArea.SecondaryTextAxis.ClearShapeProperties();
                    }

                    if (this.IsStylish) this.Legend.LegendPosition = C.LegendPositionValues.Bottom;
                    this.PlotArea.SetDataSeriesAutoAxisType();
                    this.ShowEmptyCellsAs = C.DisplayBlanksAsValues.Gap;
                    break;
                case SLStockChartType.VolumeOpenHighLowClose:
                    vType = SLDataSeriesChartType.StockChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                    iBarChartType = (int)SLDataSeriesChartType.BarChartColumnPrimary;
                    for (i = 0; i < this.PlotArea.DataSeries.Count; ++i)
                    {
                        if (i == 0)
                        {
                            this.PlotArea.DataSeries[i].ChartType = SLDataSeriesChartType.BarChartColumnPrimary;
                            if (this.IsStylish)
                            {
                                this.PlotArea.DataSeries[i].Options.Fill.SetSolidFill(A.SchemeColorValues.Accent1, 0, 0);
                                // 2.25 pt width
                                this.PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                                this.PlotArea.DataSeries[i].Options.Line.SetNoLine();
                            }

                            iBarChartType = (int)SLDataSeriesChartType.BarChartColumnPrimary;
                            this.PlotArea.UsedChartTypes[iBarChartType] = true;
                            this.PlotArea.UsedChartOptions[iBarChartType].BarDirection = C.BarDirectionValues.Column;
                            this.PlotArea.UsedChartOptions[iBarChartType].BarGrouping = C.BarGroupingValues.Clustered;
                            if (Options != null)
                            {
                                this.PlotArea.UsedChartOptions[iBarChartType].GapWidth = Options.GapWidth;
                                this.PlotArea.UsedChartOptions[iBarChartType].Overlap = Options.Overlap;
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.Fill = Options.Fill.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.Outline = Options.Border.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.Shadow = Options.Shadow.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.Glow = Options.Glow.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.EffectList.SoftEdge = Options.SoftEdge.Clone();
                                this.PlotArea.DataSeries[i].Options.ShapeProperties.Format3D = Options.Format3D.Clone();
                            }
                        }
                        else
                        {
                            this.PlotArea.DataSeries[i].ChartType = vType;
                            if (this.IsStylish)
                            {
                                // 2.25 pt width
                                this.PlotArea.DataSeries[i].Options.Line.Width = 2.25m;
                                this.PlotArea.DataSeries[i].Options.Line.CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Round;
                                this.PlotArea.DataSeries[i].Options.Line.SetNoLine();
                                this.PlotArea.DataSeries[i].Options.Line.JoinType = SLA.SLLineJoinValues.Round;
                                this.PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.None;
                            }
                        }
                    }

                    this.PlotArea.UsedChartOptions[iChartType].HasHighLowLines = true;
                    this.PlotArea.UsedChartOptions[iChartType].HasUpDownBars = true;
                    if (this.IsStylish)
                    {
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Width = 0.75m;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CapType = A.LineCapValues.Flat;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.25m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].HighLowLines.Line.JoinType = SLA.SLLineJoinValues.Round;

                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.GapWidth = 150;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Fill.SetSolidFill(A.SchemeColorValues.Light1, 0, 0);
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.Width = 0.75m;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.CapType = A.LineCapValues.Flat;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.SetSolidLine(A.SchemeColorValues.Text1, 0.35m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.UpBars.Border.JoinType = SLA.SLLineJoinValues.Round;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Fill.SetSolidFill(A.SchemeColorValues.Dark1, 0.25m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.Width = 0.75m;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.CapType = A.LineCapValues.Flat;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.SetSolidLine(A.SchemeColorValues.Text1, 0.35m, 0);
                        this.PlotArea.UsedChartOptions[iChartType].UpDownBars.DownBars.Border.JoinType = SLA.SLLineJoinValues.Round;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
                    this.PlotArea.HasSecondaryAxes = true;
                    this.PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;
                    this.PlotArea.SecondaryValueAxis.ForceAxisPosition = true;
                    this.PlotArea.SecondaryValueAxis.IsCrosses = true;
                    this.PlotArea.SecondaryValueAxis.Crosses = C.CrossesValues.Maximum;
                    this.PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
                    //this.PlotArea.SecondaryValueAxis.OtherAxisIsCrosses = true;
                    //this.PlotArea.SecondaryValueAxis.OtherAxisCrosses = C.CrossesValues.AutoZero;
                    this.PlotArea.SecondaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    this.PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
                    //this.PlotArea.SecondaryTextAxis.IsCrosses = true;
                    //this.PlotArea.SecondaryTextAxis.Crosses = C.CrossesValues.AutoZero;
                    this.PlotArea.SecondaryTextAxis.OtherAxisIsCrosses = true;
                    this.PlotArea.SecondaryTextAxis.OtherAxisCrosses = C.CrossesValues.Maximum;

                    if (this.IsStylish)
                    {
                        this.PlotArea.SecondaryValueAxis.ShowMajorGridlines = true;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.Width = 0.75m;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.CapType = A.LineCapValues.Flat;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.CompoundLineType = A.CompoundLineValues.Single;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.Alignment = A.PenAlignmentValues.Center;
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                        this.PlotArea.SecondaryValueAxis.MajorGridlines.Line.JoinType = SLA.SLLineJoinValues.Round;

                        this.PlotArea.SecondaryTextAxis.ClearShapeProperties();
                    }

                    if (this.IsStylish) this.Legend.LegendPosition = C.LegendPositionValues.Bottom;
                    this.PlotArea.SetDataSeriesAutoAxisType();
                    this.ShowEmptyCellsAs = C.DisplayBlanksAsValues.Gap;
                    break;
            }
        }

        /// <summary>
        /// Set a doughnut chart using one of the built-in doughnut chart types.
        /// </summary>
        /// <param name="ChartType">A built-in doughnut chart type.</param>
        public void SetChartType(SLDoughnutChartType ChartType)
        {
            this.SetChartType(ChartType, null);
        }

        /// <summary>
        /// Set a doughnut chart using one of the built-in doughnut chart types.
        /// </summary>
        /// <param name="ChartType">A built-in doughnut chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLDoughnutChartType ChartType, SLPieChartOptions Options)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLDoughnutChartType.Doughnut:
                    vType = SLDataSeriesChartType.DoughnutChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    this.PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                    this.PlotArea.UsedChartOptions[iChartType].HoleSize = this.IsStylish ? (byte)75 : (byte)50;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);
                    break;
                case SLDoughnutChartType.ExplodedDoughnut:
                    vType = SLDataSeriesChartType.DoughnutChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                    this.PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                    this.PlotArea.UsedChartOptions[iChartType].HoleSize = this.IsStylish ? (byte)75 : (byte)50;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Explosion = 25;
                    }
                    break;
            }
        }

        /// <summary>
        /// Set a scatter chart using one of the built-in scatter chart types.
        /// </summary>
        /// <param name="ChartType">A built-in scatter chart type.</param>
        public void SetChartType(SLScatterChartType ChartType)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLScatterChartType.ScatterWithOnlyMarkers:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Line.Width = 2.25m;
                        ds.Options.Line.SetNoLine();
                        if (this.IsStylish)
                        {
                            ds.Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            ds.Options.Marker.Size = 5;
                        }
                    }

                    this.SetPlotAreaValueAxes();
                    this.PlotArea.HasPrimaryAxes = true;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    }
                    break;
                case SLScatterChartType.ScatterWithSmoothLinesAndMarkers:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.SmoothMarker;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Smooth = true;
                        if (this.IsStylish)
                        {
                            ds.Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            ds.Options.Marker.Size = 5;
                        }
                    }

                    this.SetPlotAreaValueAxes();
                    this.PlotArea.HasPrimaryAxes = true;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    }
                    break;
                case SLScatterChartType.ScatterWithSmoothLines:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.SmoothMarker;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Smooth = true;
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;
                    }

                    this.SetPlotAreaValueAxes();
                    this.PlotArea.HasPrimaryAxes = true;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    }
                    break;
                case SLScatterChartType.ScatterWithStraightLinesAndMarkers:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    if (this.IsStylish)
                    {
                        for (int i = 0; i < this.PlotArea.DataSeries.Count; ++i)
                        {
                            this.PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            this.PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }
                    }

                    this.SetPlotAreaValueAxes();
                    this.PlotArea.HasPrimaryAxes = true;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    }
                    break;
                case SLScatterChartType.ScatterWithStraightLines:
                    vType = SLDataSeriesChartType.ScatterChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;
                    }

                    this.SetPlotAreaValueAxes();
                    this.PlotArea.HasPrimaryAxes = true;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                    }
                    break;
            }
        }

        /// <summary>
        /// Set an area chart using one of the built-in area chart types.
        /// </summary>
        /// <param name="ChartType">A built-in area chart type.</param>
        public void SetChartType(SLAreaChartType ChartType)
        {
            this.SetChartType(ChartType, null);
        }

        /// <summary>
        /// Set an area chart using one of the built-in area chart types.
        /// </summary>
        /// <param name="ChartType">A built-in area chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLAreaChartType ChartType, SLAreaChartOptions Options)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLAreaChartType.Area:
                    vType = SLDataSeriesChartType.AreaChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    break;
                case SLAreaChartType.StackedArea:
                    vType = SLDataSeriesChartType.AreaChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    break;
                case SLAreaChartType.StackedAreaMax:
                    vType = SLDataSeriesChartType.AreaChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    this.PlotArea.PrimaryValueAxis.FormatCode = "0%";
                    break;
                case SLAreaChartType.Area3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Area3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    this.PlotArea.HasDepthAxis = true;
                    this.PlotArea.DepthAxis.IsCrosses = true;
                    this.PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
                case SLAreaChartType.StackedArea3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Area3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    break;
                case SLAreaChartType.StackedAreaMax3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Area3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                    this.PlotArea.PrimaryValueAxis.FormatCode = "0%";
                    break;
            }
        }

        /// <summary>
        /// Set a line chart using one of the built-in line chart types.
        /// </summary>
        /// <param name="ChartType">A built-in line chart type.</param>
        public void SetChartType(SLLineChartType ChartType)
        {
            this.SetChartType(ChartType, null);
        }

        /// <summary>
        /// Set a line chart using one of the built-in line chart types.
        /// </summary>
        /// <param name="ChartType">A built-in line chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLLineChartType ChartType, SLLineChartOptions Options)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLLineChartType.Line:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    this.PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    this.PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.StackedLine:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    this.PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.StackedLineMax:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    this.PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    foreach (SLDataSeries ds in this.PlotArea.DataSeries)
                    {
                        ds.Options.Marker.Symbol = C.MarkerStyleValues.None;
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.LineWithMarkers:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    this.PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    this.PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    if (this.IsStylish)
                    {
                        for (int i = 0; i < this.PlotArea.DataSeries.Count; ++i)
                        {
                            this.PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            this.PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.StackedLineWithMarkers:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    this.PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    if (this.IsStylish)
                    {
                        for (int i = 0; i < this.PlotArea.DataSeries.Count; ++i)
                        {
                            this.PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            this.PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.StackedLineWithMarkersMax:
                    vType = SLDataSeriesChartType.LineChartPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                    this.PlotArea.UsedChartOptions[iChartType].Smooth = false;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    if (this.IsStylish)
                    {
                        for (int i = 0; i < this.PlotArea.DataSeries.Count; ++i)
                        {
                            this.PlotArea.DataSeries[i].Options.Marker.Symbol = C.MarkerStyleValues.Circle;
                            this.PlotArea.DataSeries[i].Options.Marker.Size = 5;
                        }
                    }

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLLineChartType.Line3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Line3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.HasDepthAxis = true;
                    this.PlotArea.DepthAxis.IsCrosses = true;
                    this.PlotArea.DepthAxis.Crosses = C.CrossesValues.AutoZero;
                    break;
            }
        }

        /// <summary>
        /// Set a column chart using one of the built-in column chart types.
        /// </summary>
        /// <param name="ChartType">A built-in column chart type.</param>
        public void SetChartType(SLColumnChartType ChartType)
        {
            this.SetChartType(ChartType, null);
        }

        /// <summary>
        /// Set a column chart using one of the built-in column chart types.
        /// </summary>
        /// <param name="ChartType">A built-in column chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLColumnChartType ChartType, SLBarChartOptions Options)
        {
            this.Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLColumnChartType.ClusteredColumn:
                    vType = SLDataSeriesChartType.BarChartColumnPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;

                    if (this.IsStylish)
                    {
                        this.PlotArea.UsedChartOptions[iChartType].Overlap = -27;
                        this.PlotArea.UsedChartOptions[iChartType].GapWidth = 219;
                    }
                    break;
                case SLColumnChartType.StackedColumn:
                    vType = SLDataSeriesChartType.BarChartColumnPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedColumnMax:
                    vType = SLDataSeriesChartType.BarChartColumnPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.ClusteredColumn3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedColumn3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedColumnMax3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.Column3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Standard;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.HasDepthAxis = true;
                    break;
                case SLColumnChartType.ClusteredCylinder:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedCylinder:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedCylinderMax:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.Cylinder3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Standard;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.HasDepthAxis = true;
                    break;
                case SLColumnChartType.ClusteredCone:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedCone:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedConeMax:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.Cone3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Standard;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.HasDepthAxis = true;
                    break;
                case SLColumnChartType.ClusteredPyramid:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedPyramid:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.StackedPyramidMax:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    break;
                case SLColumnChartType.Pyramid3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    this.RightAngleAxes = false;
                    this.Perspective = 30;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Column;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Standard;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.HasDepthAxis = true;
                    break;
            }
        }

        /// <summary>
        /// Set a bar chart using one of the built-in bar chart types.
        /// </summary>
        /// <param name="ChartType">A built-in bar chart type.</param>
        public void SetChartType(SLBarChartType ChartType)
        {
            this.SetChartType(ChartType, null);
        }

        /// <summary>
        /// Set a bar chart using one of the built-in bar chart types.
        /// </summary>
        /// <param name="ChartType">A built-in bar chart type.</param>
        /// <param name="Options">Chart customization options.</param>
        public void SetChartType(SLBarChartType ChartType, SLBarChartOptions Options)
        {
            // bar charts have their axis positions different from column charts.

            this.Is3D = SLChartTool.Is3DChart(ChartType);

            SLDataSeriesChartType vType;
            int iChartType;
            switch (ChartType)
            {
                case SLBarChartType.ClusteredBar:
                    vType = SLDataSeriesChartType.BarChartBarPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish) this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                    break;
                case SLBarChartType.StackedBar:
                    vType = SLDataSeriesChartType.BarChartBarPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish) this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                    break;
                case SLBarChartType.StackedBarMax:
                    vType = SLDataSeriesChartType.BarChartBarPrimary;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish) this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                    break;
                case SLBarChartType.ClusteredBar3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedBar3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedBarMax3D:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Box;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.ClusteredHorizontalCylinder:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalCylinder:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalCylinderMax:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cylinder;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.ClusteredHorizontalCone:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalCone:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalConeMax:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Cone;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.ClusteredHorizontalPyramid:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalPyramid:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
                case SLBarChartType.StackedHorizontalPyramidMax:
                    this.RotateX = 15;
                    this.RotateY = 20;
                    if (this.IsStylish) this.DepthPercent = 100;
                    this.RightAngleAxes = true;

                    vType = SLDataSeriesChartType.Bar3DChart;
                    this.IsCombinable = SLChartTool.IsCombinationChartFriendly(vType);
                    iChartType = (int)vType;
                    this.PlotArea.UsedChartTypes[iChartType] = true;
                    this.PlotArea.UsedChartOptions[iChartType].BarDirection = C.BarDirectionValues.Bar;
                    this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                    this.PlotArea.UsedChartOptions[iChartType].Shape = C.ShapeValues.Pyramid;
                    if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                    this.PlotArea.SetDataSeriesChartType(vType);

                    this.PlotArea.HasPrimaryAxes = true;
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;

                    if (this.IsStylish)
                    {
                        this.PlotArea.PrimaryTextAxis.MajorTickMark = C.TickMarkValues.None;
                        this.Floor.ClearShapeProperties();
                        this.Floor.Fill.SetNoFill();
                        this.Floor.Border.SetNoLine();
                    }
                    break;
            }
        }

        internal void SetPlotAreaAxes()
        {
            this.PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
            this.PlotArea.PrimaryTextAxis.AxisId = SLConstants.PrimaryAxis1;
            this.PlotArea.PrimaryTextAxis.Orientation = C.OrientationValues.MinMax;
            this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
            this.PlotArea.PrimaryTextAxis.FormatCode = SLConstants.NumberFormatGeneral;
            this.PlotArea.PrimaryTextAxis.SourceLinked = true;
            this.PlotArea.PrimaryTextAxis.HasNumberingFormat = true;
            this.PlotArea.PrimaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            this.PlotArea.PrimaryTextAxis.CrossingAxis = SLConstants.PrimaryAxis2;
            this.PlotArea.PrimaryTextAxis.IsCrosses = true;
            this.PlotArea.PrimaryTextAxis.Crosses = C.CrossesValues.AutoZero;
            this.PlotArea.PrimaryTextAxis.LabelAlignment = C.LabelAlignmentValues.Center;
            this.PlotArea.PrimaryTextAxis.LabelOffset = 100;
            this.PlotArea.PrimaryTextAxis.OtherAxisIsCrosses = true;
            this.PlotArea.PrimaryTextAxis.OtherAxisCrosses = C.CrossesValues.AutoZero;

            this.PlotArea.PrimaryValueAxis.AxisId = SLConstants.PrimaryAxis2;
            this.PlotArea.PrimaryValueAxis.Orientation = C.OrientationValues.MinMax;
            this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
            this.PlotArea.PrimaryValueAxis.ShowMajorGridlines = true;
            this.PlotArea.PrimaryValueAxis.FormatCode = SLConstants.NumberFormatGeneral;
            this.PlotArea.PrimaryValueAxis.SourceLinked = true;
            this.PlotArea.PrimaryValueAxis.HasNumberingFormat = true;
            this.PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            this.PlotArea.PrimaryValueAxis.CrossingAxis = SLConstants.PrimaryAxis1;
            this.PlotArea.PrimaryValueAxis.IsCrosses = true;
            this.PlotArea.PrimaryValueAxis.Crosses = C.CrossesValues.AutoZero;
            this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            this.PlotArea.PrimaryValueAxis.OtherAxisIsCrosses = true;
            this.PlotArea.PrimaryValueAxis.OtherAxisCrosses = C.CrossesValues.AutoZero;

            this.PlotArea.DepthAxis.AxisId = SLConstants.PrimaryAxis3;
            this.PlotArea.DepthAxis.Orientation = C.OrientationValues.MinMax;
            this.PlotArea.DepthAxis.AxisPosition = C.AxisPositionValues.Bottom;
            this.PlotArea.DepthAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            this.PlotArea.DepthAxis.CrossingAxis = SLConstants.PrimaryAxis2;

            this.PlotArea.SecondaryValueAxis.AxisId = SLConstants.SecondaryAxis2;
            this.PlotArea.SecondaryValueAxis.Orientation = C.OrientationValues.MinMax;
            this.PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;
            this.PlotArea.SecondaryValueAxis.FormatCode = SLConstants.NumberFormatGeneral;
            this.PlotArea.SecondaryValueAxis.SourceLinked = true;
            this.PlotArea.SecondaryValueAxis.HasNumberingFormat = true;
            this.PlotArea.SecondaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            this.PlotArea.SecondaryValueAxis.CrossingAxis = SLConstants.SecondaryAxis1;
            this.PlotArea.SecondaryValueAxis.IsCrosses = true;
            this.PlotArea.SecondaryValueAxis.Crosses = C.CrossesValues.Maximum;
            this.PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;

            this.PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
            this.PlotArea.SecondaryTextAxis.AxisId = SLConstants.SecondaryAxis1;
            this.PlotArea.SecondaryTextAxis.Orientation = C.OrientationValues.MinMax;
            this.PlotArea.SecondaryTextAxis.Delete = true;
            this.PlotArea.SecondaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
            this.PlotArea.SecondaryTextAxis.FormatCode = SLConstants.NumberFormatGeneral;
            this.PlotArea.SecondaryTextAxis.SourceLinked = true;
            this.PlotArea.SecondaryTextAxis.HasNumberingFormat = true;
            this.PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            this.PlotArea.SecondaryTextAxis.CrossingAxis = SLConstants.SecondaryAxis2;
            this.PlotArea.SecondaryTextAxis.LabelAlignment = C.LabelAlignmentValues.Center;
            this.PlotArea.SecondaryTextAxis.LabelOffset = 100;
            this.PlotArea.SecondaryTextAxis.OtherAxisIsCrosses = true;
            this.PlotArea.SecondaryTextAxis.OtherAxisCrosses = C.CrossesValues.Maximum;

            if (this.IsStylish)
            {
                this.PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.None;
                this.PlotArea.SecondaryValueAxis.MajorTickMark = C.TickMarkValues.None;
            }
        }

        /// <summary>
        /// This assumes SetPlotAreaAxes() is already called so fewer properties are set.
        /// </summary>
        private void SetPlotAreaValueAxes()
        {
            this.PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Value;
            this.PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;

            this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;

            // secondary axes are not set because they're dependent on what you set
            // for the chart type plotted for the data series
        }

        /// <summary>
        /// Set the position of the chart relative to the top-left of the worksheet.
        /// </summary>
        /// <param name="Top">Top position of the chart based on row index. For example, 0.5 means at the half-way point of the 1st row, 2.5 means at the half-way point of the 3rd row.</param>
        /// <param name="Left">Left position of the chart based on column index. For example, 0.5 means at the half-way point of the 1st column, 2.5 means at the half-way point of the 3rd column.</param>
        /// <param name="Bottom">Bottom position of the chart based on row index. For example, 5.5 means at the half-way point of the 6th row, 7.5 means at the half-way point of the 8th row.</param>
        /// <param name="Right">Right position of the chart based on column index. For example, 5.5 means at the half-way point of the 6th column, 7.5 means at the half-way point of the 8th column.</param>
        public void SetChartPosition(double Top, double Left, double Bottom, double Right)
        {
            double fTop = 0, fLeft = 0, fBottom = 1, fRight = 1;
            if (Top < Bottom)
            {
                fTop = Top;
                fBottom = Bottom;
            }
            else
            {
                fTop = Bottom;
                fBottom = fTop;
            }

            if (Left < Right)
            {
                fLeft = Left;
                fRight = Right;
            }
            else
            {
                fLeft = Right;
                fRight = Left;
            }

            if (fTop < 0.0) fTop = 0.0;
            if (fLeft < 0.0) fLeft = 0.0;
            if (fBottom >= SLConstants.RowLimit) fBottom = SLConstants.RowLimit;
            if (fRight >= SLConstants.ColumnLimit) fRight = SLConstants.ColumnLimit;

            this.TopPosition = fTop;
            this.LeftPosition = fLeft;
            this.BottomPosition = fBottom;
            this.RightPosition = fRight;
        }

        /// <summary>
        /// Show the chart title.
        /// </summary>
        /// <param name="Overlay">True if the title overlaps the plot area. False otherwise.</param>
        public void ShowChartTitle(bool Overlay)
        {
            this.HasTitle = true;
            this.Title.Overlay = Overlay;
        }

        /// <summary>
        /// Hide the chart title.
        /// </summary>
        public void HideChartTitle()
        {
            this.HasTitle = false;
        }

        /// <summary>
        /// Show the chart legend.
        /// </summary>
        /// <param name="Position">Position of the legend. Default is Right.</param>
        /// <param name="Overlay">True if the legend overlaps the plot area. False otherwise.</param>
        public void ShowChartLegend(C.LegendPositionValues Position, bool Overlay)
        {
            this.ShowLegend = true;
            this.Legend.LegendPosition = Position;
            this.Legend.Overlay = Overlay;
        }

        /// <summary>
        /// Hide the chart legend.
        /// </summary>
        public void HideChartLegend()
        {
            this.ShowLegend = false;
        }

        /// <summary>
        /// Get the options for a specific data series.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <returns>The data series options for the specific data series. If the index is out of bounds, a default is returned.</returns>
        public SLDataSeriesOptions GetDataSeriesOptions(int DataSeriesIndex)
        {
            SLDataSeriesOptions dso = new SLDataSeriesOptions(this.listThemeColors, this.ThrowExceptionsIfAny);

            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count)
            {
                return dso;
            }
            else
            {
                dso = this.PlotArea.DataSeries[index].Options.Clone();
                return dso;
            }
        }

        /// <summary>
        /// Set the options for a given data series.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="Options">The data series options.</param>
        public void SetDataSeriesOptions(int DataSeriesIndex, SLDataSeriesOptions Options)
        {
            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count) return;

            this.PlotArea.DataSeries[index].Options = Options.Clone();
        }

        /// <summary>
        /// Show the primary text (category/date/value) axis. This has no effect if the chart has no primary axes.
        /// </summary>
        public void ShowPrimaryTextAxis()
        {
            if (this.PlotArea.HasPrimaryAxes)
            {
                this.PlotArea.PrimaryTextAxis.Delete = false;
                this.PlotArea.PrimaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        /// Hide the primary text (category/date/value) axis. This has no effect if the chart has no primary axes.
        /// </summary>
        public void HidePrimaryTextAxis()
        {
            if (this.PlotArea.HasPrimaryAxes)
            {
                this.PlotArea.PrimaryTextAxis.Delete = true;
                this.PlotArea.PrimaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        /// Show the primary value axis. This has no effect if the chart has no primary axes.
        /// </summary>
        public void ShowPrimaryValueAxis()
        {
            if (this.PlotArea.HasPrimaryAxes)
            {
                this.PlotArea.PrimaryValueAxis.Delete = false;
                this.PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        /// Hide the primary value axis. This has no effect if the chart has no primary axes.
        /// </summary>
        public void HidePrimaryValueAxis()
        {
            if (this.PlotArea.HasPrimaryAxes)
            {
                this.PlotArea.PrimaryValueAxis.Delete = true;
                this.PlotArea.PrimaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        /// Show the depth axis. This has no effect if the chart has no depth axis (that is, not a true 3D chart).
        /// </summary>
        public void ShowDepthAxis()
        {
            if (this.PlotArea.HasDepthAxis)
            {
                this.PlotArea.DepthAxis.Delete = false;
                this.PlotArea.DepthAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        /// Hide the depth axis. This has no effect if the chart has no depth axis (that is, not a true 3D chart).
        /// </summary>
        public void HideDepthAxis()
        {
            if (this.PlotArea.HasDepthAxis)
            {
                this.PlotArea.DepthAxis.Delete = true;
                this.PlotArea.DepthAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        /// Show the secondary text (category/date/value) axis. This has no effect if the chart has no secondary axes.
        /// </summary>
        public void ShowSecondaryTextAxis()
        {
            if (this.PlotArea.HasSecondaryAxes)
            {
                if (!this.HasShownSecondaryTextAxis)
                {
                    this.HasShownSecondaryTextAxis = true;
                    this.PlotArea.SecondaryTextAxis.AxisPosition = SLChartTool.GetOppositePosition(this.PlotArea.SecondaryTextAxis.AxisPosition);
                }

                this.PlotArea.SecondaryTextAxis.Delete = false;
                this.PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        /// Hide the secondary text (category/date/value) axis. This has no effect if the chart has no secondary axes.
        /// </summary>
        public void HideSecondaryTextAxis()
        {
            if (this.PlotArea.HasSecondaryAxes)
            {
                this.PlotArea.SecondaryTextAxis.Delete = true;
                this.PlotArea.SecondaryTextAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        /// Show the secondary value axis. This has no effect if the chart has no secondary axes.
        /// </summary>
        public void ShowSecondaryValueAxis()
        {
            if (this.PlotArea.HasSecondaryAxes)
            {
                this.PlotArea.SecondaryValueAxis.Delete = false;
                this.PlotArea.SecondaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.NextTo;
            }
        }

        /// <summary>
        /// Hide the secondary value axis. This has no effect if the chart has no secondary axes.
        /// </summary>
        public void HideSecondaryValueAxis()
        {
            if (this.PlotArea.HasSecondaryAxes)
            {
                this.PlotArea.SecondaryValueAxis.Delete = true;
                this.PlotArea.SecondaryValueAxis.TickLabelPosition = C.TickLabelPositionValues.None;
            }
        }

        /// <summary>
        /// Creates an instance of SLAreaChartOptions with theme information.
        /// </summary>
        /// <returns>An SLAreaChartOptions object with theme information.</returns>
        public SLAreaChartOptions CreateAreaChartOptions()
        {
            SLAreaChartOptions aco = new SLAreaChartOptions(this.listThemeColors, this.IsStylish, this.ThrowExceptionsIfAny);
            return aco;
        }

        /// <summary>
        /// Creates an instance of SLLineChartOptions with theme information.
        /// </summary>
        /// <returns>An SLLineChartOptions object with theme information.</returns>
        public SLLineChartOptions CreateLineChartOptions()
        {
            SLLineChartOptions lco = new SLLineChartOptions(this.listThemeColors, this.IsStylish, this.ThrowExceptionsIfAny);
            return lco;
        }

        /// <summary>
        /// Creates an instance of SLPieChartOptions with theme information.
        /// </summary>
        /// <returns>An SLPieChartOptions object with theme information.</returns>
        public SLPieChartOptions CreatePieChartOptions()
        {
            SLPieChartOptions pco = new SLPieChartOptions(this.listThemeColors, this.ThrowExceptionsIfAny);
            if (this.IsStylish)
            {
                pco.Line.Width = 0.75m;
                pco.Line.CapType = A.LineCapValues.Flat;
                pco.Line.CompoundLineType = A.CompoundLineValues.Single;
                pco.Line.Alignment = A.PenAlignmentValues.Center;
                pco.Line.SetSolidLine(A.SchemeColorValues.Text1, 0.65m, 0);
                pco.Line.JoinType = SLA.SLLineJoinValues.Round;
            }
            return pco;
        }

        /// <summary>
        /// Creates an instance of SLStockChartOptions with theme information.
        /// </summary>
        /// <returns>An SLStockChartOptions object with theme information.</returns>
        public SLStockChartOptions CreateStockChartOptions()
        {
            SLStockChartOptions sco = new SLStockChartOptions(this.listThemeColors, this.IsStylish, this.ThrowExceptionsIfAny);
            return sco;
        }

        /// <summary>
        /// Plot a specific data series as a doughnut chart. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="ChartType">A built-in doughnut chart type for this specific data series.</param>
        public void PlotDataSeriesAsDoughnutChart(int DataSeriesIndex, SLDoughnutChartType ChartType)
        {
            this.PlotDataSeriesAsDoughnutChart(DataSeriesIndex, ChartType, null);
        }

        /// <summary>
        /// Plot a specific data series as a doughnut chart. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">Index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="ChartType">A built-in doughnut chart type for this specific data series.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsDoughnutChart(int DataSeriesIndex, SLDoughnutChartType ChartType, SLPieChartOptions Options)
        {
            // the original chart is not combinable
            if (!this.IsCombinable) return;

            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count) return;

            SLDataSeriesChartType vType = SLDataSeriesChartType.DoughnutChart;
            int iChartType = (int)vType;

            if (this.PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.

                // don't have to do anything if no options passed in.
                if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                this.PlotArea.UsedChartTypes[iChartType] = true;
                this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                this.PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                this.PlotArea.UsedChartOptions[iChartType].HoleSize = 50;
                if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }

            this.PlotArea.DataSeries[index].ChartType = vType;

            switch (ChartType)
            {
                case SLDoughnutChartType.Doughnut:
                    this.PlotArea.DataSeries[index].Options.iExplosion = null;
                    break;
                case SLDoughnutChartType.ExplodedDoughnut:
                    this.PlotArea.DataSeries[index].Options.Explosion = 25;
                    break;
            }
        }

        /// <summary>
        /// Plot a specific data series as a bar-of-pie chart. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        public void PlotDataSeriesAsBarOfPieChart(int DataSeriesIndex)
        {
            this.PlotDataSeriesAsOfPieChart(DataSeriesIndex, true, null);
        }

        /// <summary>
        /// Plot a specific data series as a bar-of-pie chart. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsBarOfPieChart(int DataSeriesIndex, SLPieChartOptions Options)
        {
            this.PlotDataSeriesAsOfPieChart(DataSeriesIndex, true, Options);
        }

        /// <summary>
        /// Plot a specific data series as a pie-of-pie chart. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        public void PlotDataSeriesAsPieOfPieChart(int DataSeriesIndex)
        {
            this.PlotDataSeriesAsOfPieChart(DataSeriesIndex, false, null);
        }

        /// <summary>
        /// Plot a specific data series as a pie-of-pie chart. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPieOfPieChart(int DataSeriesIndex, SLPieChartOptions Options)
        {
            this.PlotDataSeriesAsOfPieChart(DataSeriesIndex, false, Options);
        }

        private void PlotDataSeriesAsOfPieChart(int DataSeriesIndex, bool IsBarOfPie, SLPieChartOptions Options)
        {
            // the original chart is not combinable
            if (!this.IsCombinable) return;

            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count) return;

            SLDataSeriesChartType vType = IsBarOfPie ? SLDataSeriesChartType.OfPieChartBar : SLDataSeriesChartType.OfPieChartPie;
            int iChartType = (int)vType;

            if (this.PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.

                // don't have to do anything if no options passed in.
                if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                this.PlotArea.UsedChartTypes[iChartType] = true;
                this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                this.PlotArea.UsedChartOptions[iChartType].GapWidth = 100;
                this.PlotArea.UsedChartOptions[iChartType].SecondPieSize = 75;
                if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }

            this.PlotArea.DataSeries[index].ChartType = vType;
        }

        /// <summary>
        /// Plot a specific data series as a pie chart. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="IsExploded">True to explode this data series. False otherwise.</param>
        public void PlotDataSeriesAsPieChart(int DataSeriesIndex, bool IsExploded)
        {
            this.PlotDataSeriesAsPieChart(DataSeriesIndex, IsExploded, null);
        }

        /// <summary>
        /// Plot a specific data series as a pie chart. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="IsExploded">True to explode this data series. False otherwise.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPieChart(int DataSeriesIndex, bool IsExploded, SLPieChartOptions Options)
        {
            // the original chart is not combinable
            if (!this.IsCombinable) return;

            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count) return;

            SLDataSeriesChartType vType = SLDataSeriesChartType.PieChart;
            int iChartType = (int)vType;

            if (this.PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.

                // don't have to do anything if no options passed in.
                if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                this.PlotArea.UsedChartTypes[iChartType] = true;
                this.PlotArea.UsedChartOptions[iChartType].VaryColors = true;
                this.PlotArea.UsedChartOptions[iChartType].FirstSliceAngle = 0;
                if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }

            this.PlotArea.DataSeries[index].ChartType = vType;

            if (IsExploded) this.PlotArea.DataSeries[index].Options.iExplosion = null;
            else this.PlotArea.DataSeries[index].Options.Explosion = 25;
        }

        /// <summary>
        /// Plot a specific data series as a radar chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="ChartType">A built-in radar chart type for this specific data series.</param>
        public void PlotDataSeriesAsPrimaryRadarChart(int DataSeriesIndex, SLRadarChartType ChartType)
        {
            this.PlotDataSeriesAsRadarChart(DataSeriesIndex, ChartType, true);
        }

        /// <summary>
        /// Plot a specific data series as a radar chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="ChartType">A built-in radar chart type for this specific data series.</param>
        public void PlotDataSeriesAsSecondaryRadarChart(int DataSeriesIndex, SLRadarChartType ChartType)
        {
            this.PlotDataSeriesAsRadarChart(DataSeriesIndex, ChartType, false);
        }

        private void PlotDataSeriesAsRadarChart(int DataSeriesIndex, SLRadarChartType ChartType, bool IsPrimary)
        {
            // the original chart is not combinable
            if (!this.IsCombinable) return;

            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            bool bIsPrimary = IsPrimary;
            if (!this.PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                this.PlotArea.HasPrimaryAxes = true;
                this.PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
                this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                this.PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;

                this.PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }
            else if (!bIsPrimary && this.PlotArea.HasPrimaryAxes && !this.PlotArea.HasSecondaryAxes)
            {
                this.PlotArea.HasSecondaryAxes = true;
                this.PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
                this.PlotArea.SecondaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                this.PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                this.PlotArea.SecondaryTextAxis.ShowMajorGridlines = true;

                this.PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                this.PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }

            // secondary radar: cat axis is also bottom, and value axis is also left, like the primary axis.

            SLDataSeriesChartType vType = bIsPrimary ? SLDataSeriesChartType.RadarChartPrimary : SLDataSeriesChartType.RadarChartSecondary;
            int iChartType = (int)vType;

            if (this.PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.
            }
            else
            {
                this.PlotArea.UsedChartTypes[iChartType] = true;

                switch (ChartType)
                {
                    case SLRadarChartType.Radar:
                        this.PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Marker;
                        this.PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                    case SLRadarChartType.RadarWithMarkers:
                        this.PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Marker;
                        this.PlotArea.DataSeries[index].Options.Marker.vSymbol = null;
                        break;
                    case SLRadarChartType.FilledRadar:
                        this.PlotArea.UsedChartOptions[iChartType].RadarStyle = C.RadarStyleValues.Filled;
                        this.PlotArea.DataSeries[index].Options.Marker.vSymbol = null;
                        break;
                }
            }

            this.PlotArea.DataSeries[index].ChartType = vType;
        }

        /// <summary>
        /// Plot a specific data series as an area chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        public void PlotDataSeriesAsPrimaryAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            this.PlotDataSeriesAsAreaChart(DataSeriesIndex, DisplayType, null, true);
        }

        /// <summary>
        /// Plot a specific data series as an area chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPrimaryAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, SLAreaChartOptions Options)
        {
            this.PlotDataSeriesAsAreaChart(DataSeriesIndex, DisplayType, Options, true);
        }

        /// <summary>
        /// Plot a specific data series as an area chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        public void PlotDataSeriesAsSecondaryAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            this.PlotDataSeriesAsAreaChart(DataSeriesIndex, DisplayType, null, false);
        }

        /// <summary>
        /// Plot a specific data series as an area chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsSecondaryAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, SLAreaChartOptions Options)
        {
            this.PlotDataSeriesAsAreaChart(DataSeriesIndex, DisplayType, Options, false);
        }

        private void PlotDataSeriesAsAreaChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, SLAreaChartOptions Options, bool IsPrimary)
        {
            // the original chart is not combinable
            if (!this.IsCombinable) return;

            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            bool bIsPrimary = IsPrimary;
            if (!this.PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                this.PlotArea.HasPrimaryAxes = true;
                this.PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
                this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;

                this.PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
            }
            else if (!bIsPrimary && this.PlotArea.HasPrimaryAxes && !this.PlotArea.HasSecondaryAxes)
            {
                this.PlotArea.HasSecondaryAxes = true;
                this.PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
                this.PlotArea.SecondaryTextAxis.AxisPosition = this.HasShownSecondaryTextAxis ? C.AxisPositionValues.Top : C.AxisPositionValues.Bottom;
                this.PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;

                this.PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                this.PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
            }

            SLDataSeriesChartType vType = bIsPrimary ? SLDataSeriesChartType.AreaChartPrimary : SLDataSeriesChartType.AreaChartSecondary;
            int iChartType = (int)vType;

            if (this.PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.
                
                if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                this.PlotArea.UsedChartTypes[iChartType] = true;

                switch (DisplayType)
                {
                    case SLChartDataDisplayType.Normal:
                        this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                        if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                    case SLChartDataDisplayType.Stacked:
                        this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                        if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                    case SLChartDataDisplayType.StackedMax:
                        this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                        if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                        if (bIsPrimary) this.PlotArea.PrimaryValueAxis.FormatCode = "0%";
                        else this.PlotArea.SecondaryValueAxis.FormatCode = "0%";
                        break;
                }
            }

            this.PlotArea.DataSeries[index].ChartType = vType;
        }

        /// <summary>
        /// Plot a specific data series as a column chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        public void PlotDataSeriesAsPrimaryColumnChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            this.PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, null, true, false);
        }

        /// <summary>
        /// Plot a specific data series as a column chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPrimaryColumnChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, SLBarChartOptions Options)
        {
            this.PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, Options, true, false);
        }

        /// <summary>
        /// Plot a specific data series as a column chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        public void PlotDataSeriesAsSecondaryColumnChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            this.PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, null, false, false);
        }

        /// <summary>
        /// Plot a specific data series as a column chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsSecondaryColumnChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, SLBarChartOptions Options)
        {
            this.PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, Options, false, false);
        }

        /// <summary>
        /// Plot a specific data series as a bar chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        public void PlotDataSeriesAsPrimaryBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            this.PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, null, true, true);
        }

        /// <summary>
        /// Plot a specific data series as a bar chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPrimaryBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, SLBarChartOptions Options)
        {
            this.PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, Options, true, true);
        }

        /// <summary>
        /// Plot a specific data series as a bar chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        public void PlotDataSeriesAsSecondaryBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType)
        {
            this.PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, null, false, true);
        }

        /// <summary>
        /// Plot a specific data series as a bar chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsSecondaryBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, SLBarChartOptions Options)
        {
            this.PlotDataSeriesAsBarChart(DataSeriesIndex, DisplayType, Options, false, true);
        }

        private void PlotDataSeriesAsBarChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, SLBarChartOptions Options, bool IsPrimary, bool IsBar)
        {
            // the original chart is not combinable
            if (!this.IsCombinable) return;

            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            bool bIsPrimary = IsPrimary;
            if (!this.PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                this.PlotArea.HasPrimaryAxes = true;
                this.PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
                if (IsBar)
                {
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Left;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Bottom;
                }
                else
                {
                    this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                    this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                }

                this.PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }
            else if (!bIsPrimary && this.PlotArea.HasPrimaryAxes && !this.PlotArea.HasSecondaryAxes)
            {
                this.PlotArea.HasSecondaryAxes = true;
                this.PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
                if (IsBar)
                {
                    this.PlotArea.SecondaryTextAxis.AxisPosition = this.HasShownSecondaryTextAxis ? C.AxisPositionValues.Right : C.AxisPositionValues.Left;
                    this.PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Top;
                }
                else
                {
                    this.PlotArea.SecondaryTextAxis.AxisPosition = this.HasShownSecondaryTextAxis ? C.AxisPositionValues.Top : C.AxisPositionValues.Bottom;
                    this.PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;
                }

                this.PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                this.PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }

            SLDataSeriesChartType vType = SLDataSeriesChartType.BarChartBarPrimary;
            if (bIsPrimary)
            {
                if (IsBar) vType = SLDataSeriesChartType.BarChartBarPrimary;
                else vType = SLDataSeriesChartType.BarChartColumnPrimary;
            }
            else
            {
                if (IsBar) vType = SLDataSeriesChartType.BarChartBarSecondary;
                else vType = SLDataSeriesChartType.BarChartColumnSecondary;
            }

            int iChartType = (int)vType;

            if (this.PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.
                
                if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                this.PlotArea.UsedChartTypes[iChartType] = true;

                switch (DisplayType)
                {
                    case SLChartDataDisplayType.Normal:
                        this.PlotArea.UsedChartOptions[iChartType].BarDirection = IsBar ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column;
                        this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Clustered;
                        if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                    case SLChartDataDisplayType.Stacked:
                        this.PlotArea.UsedChartOptions[iChartType].BarDirection = IsBar ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column;
                        this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.Stacked;
                        this.PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                        if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                    case SLChartDataDisplayType.StackedMax:
                        this.PlotArea.UsedChartOptions[iChartType].BarDirection = IsBar ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column;
                        this.PlotArea.UsedChartOptions[iChartType].BarGrouping = C.BarGroupingValues.PercentStacked;
                        this.PlotArea.UsedChartOptions[iChartType].Overlap = 100;
                        if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
                        break;
                }
            }

            this.PlotArea.DataSeries[index].ChartType = vType;
        }

        /// <summary>
        /// Plot a specific data series as a scatter chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="ChartType">A built-in scatter chart type for this specific data series.</param>
        public void PlotDataSeriesAsPrimaryScatterChart(int DataSeriesIndex, SLScatterChartType ChartType)
        {
            this.PlotDataSeriesAsScatterChart(DataSeriesIndex, ChartType, true);
        }

        /// <summary>
        /// Plot a specific data series as a scatter chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="ChartType">A built-in scatter chart type for this specific data series.</param>
        public void PlotDataSeriesAsSecondaryScatterChart(int DataSeriesIndex, SLScatterChartType ChartType)
        {
            this.PlotDataSeriesAsScatterChart(DataSeriesIndex, ChartType, false);
        }

        private void PlotDataSeriesAsScatterChart(int DataSeriesIndex, SLScatterChartType ChartType, bool IsPrimary)
        {
            // the original chart is not combinable
            if (!this.IsCombinable) return;

            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            bool bIsPrimary = IsPrimary;
            if (!this.PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                this.PlotArea.HasPrimaryAxes = true;
                this.PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Value;
                this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                this.PlotArea.PrimaryTextAxis.ShowMajorGridlines = true;
                this.PlotArea.PrimaryValueAxis.MajorTickMark = C.TickMarkValues.Cross;

                this.PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
            }
            else if (!bIsPrimary && this.PlotArea.HasPrimaryAxes && !this.PlotArea.HasSecondaryAxes)
            {
                this.PlotArea.HasSecondaryAxes = true;
                this.PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Value;
                this.PlotArea.SecondaryTextAxis.AxisPosition = this.HasShownSecondaryTextAxis ? C.AxisPositionValues.Top : C.AxisPositionValues.Bottom;
                this.PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Left;
                this.PlotArea.SecondaryTextAxis.ShowMajorGridlines = true;

                this.PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
                this.PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.MidpointCategory;
            }

            SLDataSeriesChartType vType = bIsPrimary ? SLDataSeriesChartType.ScatterChartPrimary : SLDataSeriesChartType.ScatterChartSecondary;
            int iChartType = (int)vType;

            if (this.PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.
            }
            else
            {
                this.PlotArea.UsedChartTypes[iChartType] = true;

                switch (ChartType)
                {
                    case SLScatterChartType.ScatterWithOnlyMarkers:
                        this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                        this.PlotArea.DataSeries[index].Options.Line.Width = 2.25m;
                        this.PlotArea.DataSeries[index].Options.Line.SetNoLine();
                        break;
                    case SLScatterChartType.ScatterWithSmoothLinesAndMarkers:
                        this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.SmoothMarker;
                        this.PlotArea.DataSeries[index].Options.Smooth = true;
                        break;
                    case SLScatterChartType.ScatterWithSmoothLines:
                        this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.SmoothMarker;
                        this.PlotArea.DataSeries[index].Options.Smooth = true;
                        this.PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                    case SLScatterChartType.ScatterWithStraightLinesAndMarkers:
                        this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                        break;
                    case SLScatterChartType.ScatterWithStraightLines:
                        this.PlotArea.UsedChartOptions[iChartType].ScatterStyle = C.ScatterStyleValues.LineMarker;
                        this.PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                }
            }

            this.PlotArea.DataSeries[index].ChartType = vType;
        }

        /// <summary>
        /// Plot a specific data series as a line chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="WithMarkers">True to display markers. False otherwise.</param>
        public void PlotDataSeriesAsPrimaryLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, bool WithMarkers)
        {
            this.PlotDataSeriesAsLineChart(DataSeriesIndex, DisplayType, WithMarkers, null, true);
        }

        /// <summary>
        /// Plot a specific data series as a line chart on the primary axes. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="WithMarkers">True to display markers. False otherwise.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsPrimaryLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, bool WithMarkers, SLLineChartOptions Options)
        {
            this.PlotDataSeriesAsLineChart(DataSeriesIndex, DisplayType, WithMarkers, Options, true);
        }

        /// <summary>
        /// Plot a specific data series as a line chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="WithMarkers">True to display markers. False otherwise.</param>
        public void PlotDataSeriesAsSecondaryLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, bool WithMarkers)
        {
            this.PlotDataSeriesAsLineChart(DataSeriesIndex, DisplayType, WithMarkers, null, false);
        }

        /// <summary>
        /// Plot a specific data series as a line chart on the secondary axes. If there are no primary axes, it will be plotted on the primary axes instead. WARNING: Only weak checks done on whether the resulting combination chart is valid. Use with caution.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DisplayType">Chart display type. This corresponds to the 3 typical types in most charts: normal (or clustered), stacked and 100% stacked.</param>
        /// <param name="WithMarkers">True to display markers. False otherwise.</param>
        /// <param name="Options">Chart customization options.</param>
        public void PlotDataSeriesAsSecondaryLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, bool WithMarkers, SLLineChartOptions Options)
        {
            this.PlotDataSeriesAsLineChart(DataSeriesIndex, DisplayType, WithMarkers, Options, false);
        }

        private void PlotDataSeriesAsLineChart(int DataSeriesIndex, SLChartDataDisplayType DisplayType, bool WithMarkers, SLLineChartOptions Options, bool IsPrimary)
        {
            // the original chart is not combinable
            if (!this.IsCombinable) return;

            int index = DataSeriesIndex - 1;

            // out of bounds
            if (index < 0 || index >= this.PlotArea.DataSeries.Count) return;

            // is primary, no primary axes -> set primary axes
            // is primary, has primary axes -> do nothing
            // is secondary, no primary axes -> set primary axes, force as primary
            // is secondary, has primary axes, no secondary axes -> set secondary axes
            // is secondary, has primary axes, has secondary axes -> do nothing

            bool bIsPrimary = IsPrimary;
            if (!this.PlotArea.HasPrimaryAxes)
            {
                // no primary axes in the first place, so force primary axes
                bIsPrimary = true;
                this.PlotArea.HasPrimaryAxes = true;
                this.PlotArea.PrimaryTextAxis.AxisType = SLAxisType.Category;
                this.PlotArea.PrimaryTextAxis.AxisPosition = C.AxisPositionValues.Bottom;
                this.PlotArea.PrimaryValueAxis.AxisPosition = C.AxisPositionValues.Left;

                this.PlotArea.PrimaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                this.PlotArea.PrimaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }
            else if (!bIsPrimary && this.PlotArea.HasPrimaryAxes && !this.PlotArea.HasSecondaryAxes)
            {
                this.PlotArea.HasSecondaryAxes = true;
                this.PlotArea.SecondaryTextAxis.AxisType = SLAxisType.Category;
                this.PlotArea.SecondaryTextAxis.AxisPosition = this.HasShownSecondaryTextAxis ? C.AxisPositionValues.Top : C.AxisPositionValues.Bottom;
                this.PlotArea.SecondaryValueAxis.AxisPosition = C.AxisPositionValues.Right;

                this.PlotArea.SecondaryTextAxis.CrossBetween = C.CrossBetweenValues.Between;
                this.PlotArea.SecondaryValueAxis.CrossBetween = C.CrossBetweenValues.Between;
            }

            SLDataSeriesChartType vType = bIsPrimary ? SLDataSeriesChartType.LineChartPrimary : SLDataSeriesChartType.LineChartSecondary;
            int iChartType = (int)vType;

            if (this.PlotArea.UsedChartTypes[iChartType])
            {
                // the chart is already used.
                
                if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);
            }
            else
            {
                this.PlotArea.UsedChartTypes[iChartType] = true;

                switch (DisplayType)
                {
                    case SLChartDataDisplayType.Normal:
                        this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Standard;
                        this.PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                        if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                        if (!WithMarkers) this.PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                    case SLChartDataDisplayType.Stacked:
                        this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.Stacked;
                        this.PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                        if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                        if (!WithMarkers) this.PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                    case SLChartDataDisplayType.StackedMax:
                        this.PlotArea.UsedChartOptions[iChartType].Grouping = C.GroupingValues.PercentStacked;
                        this.PlotArea.UsedChartOptions[iChartType].ShowMarker = true;
                        if (Options != null) this.PlotArea.UsedChartOptions[iChartType].MergeOptions(Options);

                        if (!WithMarkers) this.PlotArea.DataSeries[index].Options.Marker.Symbol = C.MarkerStyleValues.None;
                        break;
                }
            }

            this.PlotArea.DataSeries[index].ChartType = vType;
        }

        /// <summary>
        /// Creates an instance of SLGroupDataLabelOptions with theme information.
        /// </summary>
        /// <returns>An SLGroupDataLabelOptions with theme information.</returns>
        public SLGroupDataLabelOptions CreateGroupDataLabelOptions()
        {
            return new SLGroupDataLabelOptions(this.listThemeColors, this.ThrowExceptionsIfAny);
        }

        /// <summary>
        /// Creates an instance of SLDataLabelOptions with theme information.
        /// </summary>
        /// <returns>An SLDataLabelOptions with theme information.</returns>
        public SLDataLabelOptions CreateDataLabelOptions()
        {
            return new SLDataLabelOptions(this.listThemeColors, this.ThrowExceptionsIfAny);
        }

        /// <summary>
        /// Set data label options to a specific data series.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="Options">Data label customization options.</param>
        public void SetGroupDataLabelOptions(int DataSeriesIndex, SLGroupDataLabelOptions Options)
        {
            // why not just return if outside of range? Because I assume you counted wrongly.
            if (DataSeriesIndex < 1) DataSeriesIndex = 1;
            if (DataSeriesIndex > this.PlotArea.DataSeries.Count) DataSeriesIndex = this.PlotArea.DataSeries.Count;
            // to get it to 0-index
            --DataSeriesIndex;

            this.PlotArea.DataSeries[DataSeriesIndex].GroupDataLabelOptions = Options.Clone();
        }

        /// <summary>
        /// Set data label options to all data series.
        /// </summary>
        /// <param name="Options">Data label customization options.</param>
        public void SetGroupDataLabelOptions(SLGroupDataLabelOptions Options)
        {
            foreach (SLDataSeries ser in this.PlotArea.DataSeries)
            {
                ser.GroupDataLabelOptions = Options.Clone();
            }
        }

        /// <summary>
        /// Set data label options to a specific data point in a specific data series.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DataPointIndex">The index of the data point. This is 1-based indexing, so it's 1 for the 1st data point, 2 for the 2nd data point and so on.</param>
        /// <param name="Options">Data label customization options.</param>
        public void SetDataLabelOptions(int DataSeriesIndex, int DataPointIndex, SLDataLabelOptions Options)
        {
            // why not just return if outside of range? Because I assume you counted wrongly.
            if (DataSeriesIndex < 1) DataSeriesIndex = 1;
            if (DataSeriesIndex > this.PlotArea.DataSeries.Count) DataSeriesIndex = this.PlotArea.DataSeries.Count;
            // to get it to 0-index
            --DataSeriesIndex;

            --DataPointIndex;
            if (DataPointIndex < 0) DataPointIndex = 0;
            this.PlotArea.DataSeries[DataSeriesIndex].DataLabelOptionsList[DataPointIndex] = Options.Clone();
        }

        /// <summary>
        /// Creates an instance of SLDataPointOptions with theme information.
        /// </summary>
        /// <returns>An SLDataPointOptions with theme information.</returns>
        public SLDataPointOptions CreateDataPointOptions()
        {
            return new SLDataPointOptions(this.listThemeColors, this.ThrowExceptionsIfAny);
        }

        /// <summary>
        /// Set data point options to a specific data point in a specific data series.
        /// </summary>
        /// <param name="DataSeriesIndex">The index of the data series. This is 1-based indexing, so it's 1 for the 1st data series, 2 for the 2nd data series and so on.</param>
        /// <param name="DataPointIndex">The index of the data point. This is 1-based indexing, so it's 1 for the 1st data point, 2 for the 2nd data point and so on.</param>
        /// <param name="Options">Data point customization options.</param>
        public void SetDataPointOptions(int DataSeriesIndex, int DataPointIndex, SLDataPointOptions Options)
        {
            // why not just return if outside of range? Because I assume you counted wrongly.
            if (DataSeriesIndex < 1) DataSeriesIndex = 1;
            if (DataSeriesIndex > this.PlotArea.DataSeries.Count) DataSeriesIndex = this.PlotArea.DataSeries.Count;
            // to get it to 0-index
            --DataSeriesIndex;

            --DataPointIndex;
            if (DataPointIndex < 0) DataPointIndex = 0;
            this.PlotArea.DataSeries[DataSeriesIndex].DataPointOptionsList[DataPointIndex] = Options.Clone();
        }

        internal C.ChartSpace ToChartSpace(ref ChartPart chartp)
        {
            ImagePart imgp;

            C.ChartSpace cs = new C.ChartSpace();
            cs.AddNamespaceDeclaration("c", SLConstants.NamespaceC);
            cs.AddNamespaceDeclaration("a", SLConstants.NamespaceA);
            cs.AddNamespaceDeclaration("r", SLConstants.NamespaceRelationships);

            cs.Date1904 = new C.Date1904() { Val = this.Date1904 };
            
            cs.EditingLanguage = new C.EditingLanguage();
            cs.EditingLanguage.Val = System.Globalization.CultureInfo.CurrentCulture.Name;

            cs.RoundedCorners = new C.RoundedCorners() { Val = this.RoundedCorners };

            AlternateContent altcontent = new AlternateContent();
            altcontent.AddNamespaceDeclaration("mc", SLConstants.NamespaceMc);

            AlternateContentChoice altcontentchoice = new AlternateContentChoice() { Requires = "c14" };
            altcontentchoice.AddNamespaceDeclaration("c14", SLConstants.NamespaceC14);
            // why +100? I don't know... ask Microsoft. But there are 48 styles. Even with the
            // advanced "+100" version, it's 96 total. It's a byte, with 256 possibilities.
            // As of this writing, Excel 2013 is rumoured to dispense away with this chart styling.
            // So maybe all this is moot anyway...
            altcontentchoice.Append(new C14.Style() { Val = (byte)(this.ChartStyle + 100) });
            altcontent.Append(altcontentchoice);

            AlternateContentFallback altcontentfallback = new AlternateContentFallback();
            altcontentfallback.Append(new C.Style() { Val = (byte)this.ChartStyle });
            altcontent.Append(altcontentfallback);

            cs.Append(altcontent);

            C.Chart chart = new C.Chart();

            if (this.HasView3D)
            {
                chart.View3D = new C.View3D();
                if (this.RotateX != null) chart.View3D.RotateX = new C.RotateX() { Val = this.RotateX.Value };
                if (this.HeightPercent != null) chart.View3D.HeightPercent = new C.HeightPercent() { Val = this.HeightPercent.Value };
                if (this.RotateY != null) chart.View3D.RotateY = new C.RotateY() { Val = this.RotateY.Value };
                if (this.DepthPercent != null) chart.View3D.DepthPercent = new C.DepthPercent() { Val = this.DepthPercent };
                if (this.RightAngleAxes != null) chart.View3D.RightAngleAxes = new C.RightAngleAxes() { Val = this.RightAngleAxes.Value };
                if (this.Perspective != null) chart.View3D.Perspective = new C.Perspective() { Val = this.Perspective.Value };
            }

            if (this.HasTitle)
            {
                if (this.Title.Fill.Type == SLA.SLFillType.BlipFill)
                {
                    imgp = chartp.AddImagePart(SLA.SLDrawingTool.GetImagePartType(this.Title.Fill.BlipFileName));
                    using (FileStream fs = new FileStream(this.Title.Fill.BlipFileName, FileMode.Open))
                    {
                        imgp.FeedData(fs);
                    }
                    this.Title.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                }
                chart.Title = this.Title.ToTitle(IsStylish);
            }
            else
            {
                chart.AutoTitleDeleted = new C.AutoTitleDeleted() { Val = true };
            }

            if (this.Is3D)
            {
                chart.Floor = new C.Floor();
                chart.Floor.Thickness = new C.Thickness() { Val = this.Floor.Thickness };
                if (this.Floor.ShapeProperties.HasShapeProperties)
                {
                    if (this.Floor.Fill.Type == SLA.SLFillType.BlipFill)
                    {
                        imgp = chartp.AddImagePart(SLA.SLDrawingTool.GetImagePartType(this.Floor.Fill.BlipFileName));
                        using (FileStream fs = new FileStream(this.Floor.Fill.BlipFileName, FileMode.Open))
                        {
                            imgp.FeedData(fs);
                        }
                        this.Floor.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                    }
                    chart.Floor.ShapeProperties = this.Floor.ShapeProperties.ToCShapeProperties(IsStylish);
                }

                chart.SideWall = new C.SideWall();
                chart.SideWall.Thickness = new C.Thickness() { Val = this.SideWall.Thickness };
                if (this.SideWall.ShapeProperties.HasShapeProperties)
                {
                    if (this.SideWall.Fill.Type == SLA.SLFillType.BlipFill)
                    {
                        imgp = chartp.AddImagePart(SLA.SLDrawingTool.GetImagePartType(this.SideWall.Fill.BlipFileName));
                        using (FileStream fs = new FileStream(this.SideWall.Fill.BlipFileName, FileMode.Open))
                        {
                            imgp.FeedData(fs);
                        }
                        this.SideWall.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                    }
                    chart.SideWall.ShapeProperties = this.SideWall.ShapeProperties.ToCShapeProperties(this.IsStylish);
                }

                chart.BackWall = new C.BackWall();
                chart.BackWall.Thickness = new C.Thickness() { Val = this.BackWall.Thickness };
                if (this.BackWall.ShapeProperties.HasShapeProperties)
                {
                    if (this.BackWall.Fill.Type == SLA.SLFillType.BlipFill)
                    {
                        imgp = chartp.AddImagePart(SLA.SLDrawingTool.GetImagePartType(this.BackWall.Fill.BlipFileName));
                        using (FileStream fs = new FileStream(this.BackWall.Fill.BlipFileName, FileMode.Open))
                        {
                            imgp.FeedData(fs);
                        }
                        this.BackWall.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                    }
                    chart.BackWall.ShapeProperties = this.BackWall.ShapeProperties.ToCShapeProperties(this.IsStylish);
                }
            }

            if (this.PlotArea.Fill.Type == SLA.SLFillType.BlipFill)
            {
                imgp = chartp.AddImagePart(SLA.SLDrawingTool.GetImagePartType(this.PlotArea.Fill.BlipFileName));
                using (FileStream fs = new FileStream(this.PlotArea.Fill.BlipFileName, FileMode.Open))
                {
                    imgp.FeedData(fs);
                }
                this.PlotArea.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
            }
            chart.PlotArea = this.PlotArea.ToPlotArea(this.IsStylish);

            if (this.ShowLegend)
            {
                if (this.Legend.Fill.Type == SLA.SLFillType.BlipFill)
                {
                    imgp = chartp.AddImagePart(SLA.SLDrawingTool.GetImagePartType(this.Legend.Fill.BlipFileName));
                    using (FileStream fs = new FileStream(this.Legend.Fill.BlipFileName, FileMode.Open))
                    {
                        imgp.FeedData(fs);
                    }
                    this.Legend.Fill.BlipRelationshipID = chartp.GetIdOfPart(imgp);
                }
                chart.Legend = this.Legend.ToLegend(this.IsStylish);
            }

            chart.PlotVisibleOnly = new C.PlotVisibleOnly() { Val = !this.ShowHiddenData };

            chart.DisplayBlanksAs = new C.DisplayBlanksAs() { Val = this.ShowEmptyCellsAs };

            chart.ShowDataLabelsOverMaximum = new C.ShowDataLabelsOverMaximum() { Val = this.ShowDataLabelsOverMaximum };

            cs.Append(chart);

            if (this.ShapeProperties.HasShapeProperties)
            {
                cs.Append(this.ShapeProperties.ToCShapeProperties(IsStylish));
            }

            return cs;
        }

        internal SLChart Clone()
        {
            SLChart chart = new SLChart();
            chart.listThemeColors = new List<System.Drawing.Color>();
            for (int i = 0; i < this.listThemeColors.Count; ++i)
            {
                chart.listThemeColors.Add(this.listThemeColors[i]);
            }

            chart.Date1904 = this.Date1904;
            chart.IsStylish = this.IsStylish;
            chart.RoundedCorners = this.RoundedCorners;
            chart.IsCombinable = this.IsCombinable;

            chart.TopPosition = this.TopPosition;
            chart.LeftPosition = this.LeftPosition;
            chart.BottomPosition = this.BottomPosition;
            chart.RightPosition = this.RightPosition;
            chart.WorksheetName = this.WorksheetName;
            chart.RowsAsDataSeries = this.RowsAsDataSeries;
            chart.ShowHiddenData = this.ShowHiddenData;
            chart.ShowDataLabelsOverMaximum = this.ShowDataLabelsOverMaximum;

            chart.StartRowIndex = this.StartRowIndex;
            chart.StartColumnIndex = this.StartColumnIndex;
            chart.EndRowIndex = this.EndRowIndex;
            chart.EndColumnIndex = this.EndColumnIndex;

            chart.ChartStyle = this.ChartStyle;
            chart.ShowEmptyCellsAs = this.ShowEmptyCellsAs;
            chart.RotateX = this.RotateX;
            chart.HeightPercent = this.HeightPercent;
            chart.RotateY = this.RotateY;
            chart.DepthPercent = this.DepthPercent;
            chart.RightAngleAxes = this.RightAngleAxes;
            chart.Perspective = this.Perspective;
            chart.ChartName = this.ChartName;

            chart.HasTitle = this.HasTitle;
            chart.Title = this.Title.Clone();
            
            chart.Is3D = this.Is3D;

            chart.Floor = this.Floor.Clone();
            chart.SideWall = this.SideWall.Clone();
            chart.BackWall = this.BackWall.Clone();
            chart.PlotArea = this.PlotArea.Clone();
            chart.HasShownSecondaryTextAxis = this.HasShownSecondaryTextAxis;
            chart.ShowLegend = this.ShowLegend;
            chart.Legend = this.Legend.Clone();
            chart.ShapeProperties = this.ShapeProperties.Clone();

            return chart;
        }
    }
}
