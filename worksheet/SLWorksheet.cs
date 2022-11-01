using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace SpreadsheetLight
{
    internal class SLWorksheet
    {
        internal bool ForceCustomRowColumnDimensionsSplitting { get; set; }

        /// <summary>
        /// The default is A1. Not going to get the active cell from the worksheet.
        /// This is purely for setting (not getting). If not A1, then we'll do stuff.
        /// </summary>
        internal SLCellPoint ActiveCell { get; set; }
        
        internal List<SLSheetView> SheetViews { get; set; }

        internal bool IsDoubleColumnWidth { get; set; }
        internal SLSheetFormatProperties SheetFormatProperties { get; set; }

        // For posterity, here's a note about styles:
        // It's column style, then row style, then cell style. In increasing priority.
        // So if a cell has the default style, then we check if the row it belongs to
        // has a style. If the row also has no style (haha), then we check if the column
        // the cell belongs to has a style. If all are default, then the cell is truly
        // without any fashion sense.
        // This seems to be Excel's way of cascading styles, so we follow.

        internal Dictionary<int, SLRowProperties> RowProperties { get; set; }
        internal Dictionary<int, SLColumnProperties> ColumnProperties { get; set; }
        //...internal Dictionary<SLCellPoint, SLCell> Cells { get; set; }
        internal SLCellWarehouse CellWarehouse { get; set; }

        internal Dictionary<uint, SLSharedCellFormula> SharedCellFormulas { get; set; }

        // note that this doesn't mean that the worksheet is protected,
        // just that the SheetProtection SDK class is present.
        internal bool HasSheetProtection;
        internal SLSheetProtection SheetProtection { get; set; }

        internal bool HasAutoFilter;
        internal SLAutoFilter AutoFilter { get; set; }

        internal List<SLMergeCell> MergeCells { get; set; }

        internal List<SLConditionalFormatting> ConditionalFormattings { get; set; }
        internal List<SLConditionalFormatting2010> ConditionalFormattings2010 { get; set; }
        internal List<SLConditionalFormatOrder> ConditionalFormatOrder { get; set; }

        internal List<SLDataValidation> DataValidations { get; set; }
        internal bool DataValidationDisablePrompts { get; set; }
        internal uint? DataValidationXWindow { get; set; }
        internal uint? DataValidationYWindow { get; set; }

        internal List<SLHyperlink> Hyperlinks { get; set; }

        internal SLPageSettings PageSettings { get; set; }

        internal Dictionary<int, SLBreak> RowBreaks { get; set; }
        internal Dictionary<int, SLBreak> ColumnBreaks { get; set; }

        // use the reference ID of the Drawing class directly
        internal string DrawingId { get; set; }

        internal uint NextWorksheetDrawingId { get; set; }

        internal List<Drawing.SLPicture> Pictures { get; set; }

        internal List<Charts.SLChart> Charts { get; set; }

        internal bool ToAppendBackgroundPicture { get; set; }
        internal string BackgroundPictureId { get; set; }
        /// <summary>
        /// if null, then don't have to do anything
        /// </summary>
        internal bool? BackgroundPictureDataIsInFile { get; set; }
        internal string BackgroundPictureFileName { get; set; }
        internal byte[] BackgroundPictureByteData { get; set; }
        internal ImagePartType BackgroundPictureImagePartType { get; set; }

        // for cell comments
        internal string LegacyDrawingId { get; set; }
        internal List<string> Authors { get; set; }
        internal Dictionary<SLCellPoint, SLComment> Comments { get; set; }

        internal List<SLTable> Tables { get; set; }

        internal List<SLSparklineGroup> SparklineGroups { get; set; }

        internal SLWorksheet(List<System.Drawing.Color> ThemeColors, List<System.Drawing.Color> IndexedColors, double ThemeDefaultColumnWidth, long ThemeDefaultColumnWidthInEMU, int MaxDigitWidth, List<double> ColumnStepSize, double CalculatedDefaultRowHeight)
        {
            this.ForceCustomRowColumnDimensionsSplitting = false;

            this.ActiveCell = new SLCellPoint(1, 1);

            this.SheetViews = new List<SLSheetView>();

            this.IsDoubleColumnWidth = false;
            this.SheetFormatProperties = new SLSheetFormatProperties(ThemeDefaultColumnWidth, ThemeDefaultColumnWidthInEMU, MaxDigitWidth, ColumnStepSize, CalculatedDefaultRowHeight);

            this.RowProperties = new Dictionary<int, SLRowProperties>();
            this.ColumnProperties = new Dictionary<int, SLColumnProperties>();
            //...this.Cells = new Dictionary<SLCellPoint, SLCell>();
            this.CellWarehouse = new SLCellWarehouse();

            this.SharedCellFormulas = new Dictionary<uint, SLSharedCellFormula>();

            this.HasSheetProtection = false;
            this.SheetProtection = new SLSheetProtection();

            this.HasAutoFilter = false;
            this.AutoFilter = new SLAutoFilter();

            this.MergeCells = new List<SLMergeCell>();

            this.ConditionalFormattings = new List<SLConditionalFormatting>();
            this.ConditionalFormattings2010 = new List<SLConditionalFormatting2010>();
            this.ConditionalFormatOrder = new List<SLConditionalFormatOrder>();

            this.DataValidations = new List<SLDataValidation>();
            this.DataValidationDisablePrompts = false;
            this.DataValidationXWindow = null;
            this.DataValidationYWindow = null;

            this.Hyperlinks = new List<SLHyperlink>();

            this.PageSettings = new SLPageSettings(ThemeColors, IndexedColors);

            this.RowBreaks = new Dictionary<int, SLBreak>();
            this.ColumnBreaks = new Dictionary<int, SLBreak>();

            this.DrawingId = string.Empty;
            this.NextWorksheetDrawingId = 2;
            this.Pictures = new List<Drawing.SLPicture>();
            this.Charts = new List<Charts.SLChart>();

            this.InitializeBackgroundPictureStuff();

            this.LegacyDrawingId = string.Empty;
            this.Authors = new List<string>();
            this.Comments = new Dictionary<SLCellPoint, SLComment>();

            this.Tables = new List<SLTable>();

            this.SparklineGroups = new List<SLSparklineGroup>();
        }

        internal void InitializeBackgroundPictureStuff()
        {
            this.BackgroundPictureId = string.Empty;
            this.BackgroundPictureDataIsInFile = null;
            this.BackgroundPictureFileName = string.Empty;
            this.BackgroundPictureByteData = new byte[1];
            this.BackgroundPictureImagePartType = ImagePartType.Bmp;
        }

        internal void ToggleCustomRowColumnDimension(bool IsCustom)
        {
            this.SheetFormatProperties.HasDefaultColumnWidth = IsCustom;
            if (IsCustom)
            {
                this.SheetFormatProperties.CustomHeight = IsCustom;
            }
            else
            {
                // default is false
                this.SheetFormatProperties.CustomHeight = null;
            }
        }

        internal void RefreshSparklineGroups()
        {
            for (int i = this.SparklineGroups.Count - 1; i >= 0; --i)
            {
                // in case the group has no sparklines
                if (this.SparklineGroups[i].Sparklines.Count == 0)
                {
                    this.SparklineGroups.RemoveAt(i);
                }
            }
        }
    }
}
