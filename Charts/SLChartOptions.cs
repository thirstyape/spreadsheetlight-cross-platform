using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight.Charts
{
    internal class SLChartOptions
    {
        internal C.BarDirectionValues BarDirection { get; set; }
        internal C.BarGroupingValues BarGrouping { get; set; }

        internal bool? VaryColors { get; set; }

        private ushort iGapWidth;
        internal ushort GapWidth
        {
            get { return iGapWidth; }
            set
            {
                iGapWidth = value;
                if (iGapWidth > 500) iGapWidth = 500;
            }
        }

        private ushort iGapDepth;
        internal ushort GapDepth
        {
            get { return iGapDepth; }
            set
            {
                iGapDepth = value;
                if (iGapDepth > 500) iGapDepth = 500;
            }
        }

        private sbyte byOverlap;
        internal sbyte Overlap
        {
            get { return byOverlap; }
            set
            {
                byOverlap = value;
                if (byOverlap < -100) byOverlap = -100;
                if (byOverlap > 100) byOverlap = 100;
            }
        }

        internal C.ShapeValues Shape { get; set; }

        internal C.GroupingValues Grouping { get; set; }

        internal bool ShowMarker { get; set; }
        internal bool Smooth { get; set; }

        private ushort iFirstSliceAngle;
        internal ushort FirstSliceAngle
        {
            get { return iFirstSliceAngle; }
            set
            {
                iFirstSliceAngle = value;
                if (iFirstSliceAngle > 360) iFirstSliceAngle = 360;
            }
        }

        private byte byHoleSize;
        internal byte HoleSize
        {
            get { return byHoleSize; }
            set
            {
                byHoleSize = value;
                if (byHoleSize < 10) byHoleSize = 10;
                if (byHoleSize > 90) byHoleSize = 90;
            }
        }

        internal bool HasSplit;
        internal C.SplitValues SplitType { get; set; }
        internal double SplitPosition { get; set; }
        internal List<int> SecondPiePoints { get; set; }

        private ushort iSecondPieSize;
        internal ushort SecondPieSize
        {
            get { return iSecondPieSize; }
            set
            {
                iSecondPieSize = value;
                if (iSecondPieSize < 5) iSecondPieSize = 5;
                if (iSecondPieSize > 200) iSecondPieSize = 200;
            }
        }

        // for the series line of of-pie charts
        internal SLA.SLShapeProperties SeriesLinesShapeProperties;

        internal C.ScatterStyleValues ScatterStyle { get; set; }

        internal bool? bWireframe;
        internal bool Wireframe
        {
            get { return bWireframe ?? true; }
            set { bWireframe = value; }
        }

        internal C.RadarStyleValues RadarStyle { get; set; }

        internal bool Bubble3D { get; set; }

        private uint iBubbleScale;
        internal uint BubbleScale
        {
            get { return iBubbleScale; }
            set
            {
                iBubbleScale = value;
                if (iBubbleScale > 300) iBubbleScale = 300;
            }
        }

        internal bool ShowNegativeBubbles { get; set; }
        internal C.SizeRepresentsValues SizeRepresents { get; set; }

        internal bool HasDropLines;
        internal SLDropLines DropLines { get; set; }

        internal bool HasHighLowLines;
        internal SLHighLowLines HighLowLines { get; set; }

        internal bool HasUpDownBars;
        internal SLUpDownBars UpDownBars { get; set; }

        internal SLChartOptions(List<System.Drawing.Color> ThemeColors, bool IsStylish, bool ThrowExceptionsIfAny)
        {
            this.BarDirection = C.BarDirectionValues.Bar;
            this.BarGrouping = C.BarGroupingValues.Standard;
            this.VaryColors = null;
            this.GapWidth = 150;
            this.GapDepth = 150;
            this.Overlap = 0;
            this.Shape = C.ShapeValues.Box;
            this.Grouping = C.GroupingValues.Standard;
            this.ShowMarker = true;
            this.Smooth = false;
            this.FirstSliceAngle = 0;
            this.HoleSize = 10;
            this.HasSplit = false;
            this.SplitType = C.SplitValues.Position;
            this.SplitPosition = 0;
            this.SecondPiePoints = new List<int>();
            this.SecondPieSize = 75;
            this.SeriesLinesShapeProperties = new SLA.SLShapeProperties(ThemeColors, ThrowExceptionsIfAny);
            this.ScatterStyle = C.ScatterStyleValues.Line;
            this.bWireframe = null;
            this.RadarStyle = C.RadarStyleValues.Standard;
            this.Bubble3D = true;
            this.BubbleScale = 100;
            this.ShowNegativeBubbles = true;
            this.SizeRepresents = C.SizeRepresentsValues.Area;
            this.HasDropLines = false;
            this.DropLines = new SLDropLines(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.HasHighLowLines = false;
            this.HighLowLines = new SLHighLowLines(ThemeColors, IsStylish, ThrowExceptionsIfAny);
            this.HasUpDownBars = false;
            this.UpDownBars = new SLUpDownBars(ThemeColors, IsStylish, ThrowExceptionsIfAny);
        }

        internal void MergeOptions(SLBarChartOptions bco)
        {
            this.GapWidth = bco.GapWidth;
            this.GapDepth = bco.GapDepth;
            this.Overlap = bco.Overlap;
        }

        internal void MergeOptions(SLLineChartOptions lco)
        {
            this.GapDepth = lco.GapDepth;
            this.HasDropLines = lco.HasDropLines;
            this.DropLines = lco.DropLines.Clone();
            this.HasHighLowLines = lco.HasHighLowLines;
            this.HighLowLines = lco.HighLowLines.Clone();
            this.HasUpDownBars = lco.HasUpDownBars;
            this.UpDownBars = lco.UpDownBars.Clone();
            this.Smooth = lco.Smooth;
        }

        internal void MergeOptions(SLPieChartOptions pco)
        {
            this.VaryColors = pco.VaryColors;
            this.FirstSliceAngle = pco.FirstSliceAngle;
            this.HoleSize = pco.HoleSize;
            this.GapWidth = pco.GapWidth;
            this.HasSplit = pco.HasSplit;
            this.SplitType = pco.SplitType;
            this.SplitPosition = pco.SplitPosition;

            this.SecondPiePoints.Clear();
            foreach (int i in pco.SecondPiePoints)
            {
                this.SecondPiePoints.Add(i);
            }
            this.SecondPiePoints.Sort();

            this.SecondPieSize = pco.SecondPieSize;

            this.SeriesLinesShapeProperties = pco.ShapeProperties.Clone();
        }

        internal void MergeOptions(SLAreaChartOptions aco)
        {
            this.HasDropLines = aco.HasDropLines;
            this.DropLines = aco.DropLines.Clone();
            this.GapDepth = aco.GapDepth;
        }

        internal void MergeOptions(SLBubbleChartOptions bco)
        {
            this.Bubble3D = bco.Bubble3D;
            this.BubbleScale = bco.BubbleScale;
            this.ShowNegativeBubbles = bco.ShowNegativeBubbles;
            this.SizeRepresents = bco.SizeRepresents;
        }

        internal void MergeOptions(SLStockChartOptions sco)
        {
            this.HasDropLines = sco.HasDropLines;
            this.DropLines = sco.DropLines.Clone();
            this.HasHighLowLines = sco.HasHighLowLines;
            this.HighLowLines = sco.HighLowLines.Clone();
            this.HasUpDownBars = sco.HasUpDownBars;
            this.UpDownBars = sco.UpDownBars.Clone();
        }

        internal SLChartOptions Clone()
        {
            SLChartOptions co = new SLChartOptions(this.SeriesLinesShapeProperties.listThemeColors, false, this.SeriesLinesShapeProperties.ThrowExceptionsIfAny);
            co.BarDirection = this.BarDirection;
            co.BarGrouping = this.BarGrouping;
            co.VaryColors = this.VaryColors;
            co.iGapWidth = this.iGapWidth;
            co.iGapDepth = this.iGapDepth;
            co.byOverlap = this.byOverlap;
            co.Shape = this.Shape;
            co.Grouping = this.Grouping;
            co.ShowMarker = this.ShowMarker;
            co.Smooth = this.Smooth;
            co.iFirstSliceAngle = this.iFirstSliceAngle;
            co.byHoleSize = this.byHoleSize;
            co.HasSplit = this.HasSplit;
            co.SplitType = this.SplitType;
            co.SplitPosition = this.SplitPosition;

            co.SecondPiePoints = new List<int>();
            for (int i = 0; i < this.SecondPiePoints.Count; ++i)
            {
                co.SecondPiePoints.Add(this.SecondPiePoints[i]);
            }

            co.iSecondPieSize = this.iSecondPieSize;
            co.SeriesLinesShapeProperties = this.SeriesLinesShapeProperties.Clone();
            co.ScatterStyle = this.ScatterStyle;
            co.bWireframe = this.bWireframe;
            co.RadarStyle = this.RadarStyle;
            co.Bubble3D = this.Bubble3D;
            co.iBubbleScale = this.iBubbleScale;
            co.ShowNegativeBubbles = this.ShowNegativeBubbles;
            co.SizeRepresents = this.SizeRepresents;

            co.HasDropLines = this.HasDropLines;
            co.DropLines = this.DropLines.Clone();
            co.HasHighLowLines = this.HasHighLowLines;
            co.HighLowLines = this.HighLowLines.Clone();
            co.HasUpDownBars = this.HasUpDownBars;
            co.UpDownBars = this.UpDownBars.Clone();

            return co;
        }
    }
}
