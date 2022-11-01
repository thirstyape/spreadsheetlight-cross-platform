using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        //TODO

        //public bool InsertPivotTable(SLPivotTable PivotTable, string CellReference)
        //{
        //    int iRowIndex = -1;
        //    int iColumnIndex = -1;
        //    if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
        //    {
        //        return false;
        //    }

        //    return InsertPivotTable(PivotTable, gsSelectedWorksheetName, iRowIndex, iColumnIndex, false);
        //}

        //public bool InsertPivotTable(SLPivotTable PivotTable, int RowIndex, int ColumnIndex)
        //{
        //    return InsertPivotTable(PivotTable, gsSelectedWorksheetName, RowIndex, ColumnIndex, false);
        //}

        //public bool InsertPivotTable(SLPivotTable PivotTable, string NewWorksheetName, string CellReference)
        //{
        //    int iRowIndex = -1;
        //    int iColumnIndex = -1;
        //    if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
        //    {
        //        return false;
        //    }

        //    return InsertPivotTable(PivotTable, NewWorksheetName, iRowIndex, iColumnIndex, false);
        //}

        //public bool InsertPivotTable(SLPivotTable PivotTable, string NewWorksheetName, int RowIndex, int ColumnIndex)
        //{
        //    return InsertPivotTable(PivotTable, NewWorksheetName, RowIndex, ColumnIndex, false);
        //}

        //public bool InsertPivotTable(SLPivotTable PivotTable, string NewWorksheetName, string CellReference, bool SelectNewWorksheet)
        //{
        //    int iRowIndex = -1;
        //    int iColumnIndex = -1;
        //    if (!SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
        //    {
        //        return false;
        //    }

        //    return InsertPivotTable(PivotTable, NewWorksheetName, iRowIndex, iColumnIndex, false);
        //}

        //public bool InsertPivotTable(SLPivotTable PivotTable, string NewWorksheetName, int RowIndex, int ColumnIndex, bool SelectNewWorksheet)
        //{
        //    // the upper limits are checked for consistency (well, for fun basically).
        //    // In actuality, if the pivot table is set at the bottom/right of the worksheet,
        //    // there's nothing to show anyway.
        //    if (RowIndex < 1) RowIndex = 1;
        //    if (RowIndex > SLConstants.RowLimit) RowIndex = SLConstants.RowLimit;
        //    if (ColumnIndex < 1) ColumnIndex = 1;
        //    if (ColumnIndex > SLConstants.ColumnLimit) ColumnIndex = SLConstants.ColumnLimit;

        //    if (!PivotTable.IsValid) return false;

        //    if (!SLTool.CheckSheetChartName(NewWorksheetName))
        //    {
        //        return false;
        //    }

        //    // Get cell data first
        //    Dictionary<SLCellPoint, SLCell> cells = new Dictionary<SLCellPoint, SLCell>();

        //    if (PivotTable.IsDataSourceTable)
        //    {
        //        //TODO
        //    }
        //    else
        //    {
        //        // else data source is from worksheet, and we're taking that as the current worksheet.
        //        List<SLCellPoint> currentpts = slws.Cells.Keys.ToList<SLCellPoint>();
        //        foreach (SLCellPoint currentpt in currentpts)
        //        {
        //            if (PivotTable.DataRange.StartRowIndex <= currentpt.RowIndex
        //                && currentpt.RowIndex <= PivotTable.DataRange.EndRowIndex
        //                && PivotTable.DataRange.StartColumnIndex <= currentpt.ColumnIndex
        //                && currentpt.ColumnIndex <= PivotTable.DataRange.EndColumnIndex)
        //            {
        //                cells[currentpt] = slws.Cells[currentpt].Clone();
        //            }
        //        }
        //    }

        //    // keep the current worksheet name in case we have to select it again
        //    string sCurrentWorksheetName = gsSelectedWorksheetName;

        //    string sNewWorksheetName = NewWorksheetName;

        //    HashSet<string> hsSheetNames = new HashSet<string>();
        //    foreach (SLSheet sheet in slwb.Sheets)
        //    {
        //        // there shouldn't be name collisions, so we're not checking
        //        hsSheetNames.Add(sheet.Name);
        //    }

        //    if (hsSheetNames.Contains(sNewWorksheetName))
        //    {
        //        if (!sNewWorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
        //        {
        //            // an existing sheet name is used, but is not the current worksheet
        //            // So we must come up with a new name.
        //            int index = 1;
        //            sNewWorksheetName = string.Format("Sheet{0}", index);
        //            while (hsSheetNames.Contains(sNewWorksheetName))
        //            {
        //                ++index;
        //                sNewWorksheetName = string.Format("Sheet{0}", index);
        //            }

        //            // this automatically selects the worksheet too.
        //            AddWorksheet(sNewWorksheetName);
        //        }
        //        // else we just use the current worksheet.
        //        // Yup, that's it. Don't have to do anything else.
        //    }
        //    else
        //    {
        //        // this automatically selects the worksheet too.
        //        AddWorksheet(sNewWorksheetName);
        //    }

        //    int i, j, iSharedStringIndex;
        //    SLCell cell;
        //    SLCellPoint pt;
        //    SLRstType rst = new SLRstType();
        //    string sPlainString = string.Empty;
        //    Dictionary<int, string> dictSharedStringSimplified = new Dictionary<int, string>();

        //    HashSet<string> hsSharedItemString = new HashSet<string>();
        //    HashSet<double> hsSharedItemNumber = new HashSet<double>();
        //    HashSet<bool> hsSharedItemBoolean = new HashSet<bool>();
        //    // missing, error, datetime?

        //    PivotTableCacheDefinitionPart cachedef = wbp.AddNewPart<PivotTableCacheDefinitionPart>();

        //    #region Content for PivotCacheDefinition
        //    SLPivotCacheDefinition pcd = new SLPivotCacheDefinition();

        //    SLCellPointRange datarange = new SLCellPointRange();

        //    if (PivotTable.IsDataSourceTable)
        //    {
        //        //TODO source is a Table
        //    }
        //    else
        //    {
        //        pcd.CacheSource.IsWorksheetSource = true;
        //        pcd.CacheSource.WorksheetSourceReference = SLTool.TranslateCellPointRangeToReference(PivotTable.Location.Reference);
        //        pcd.CacheSource.WorksheetSourceSheet = gsSelectedWorksheetName;

        //        datarange.StartRowIndex = PivotTable.Location.Reference.StartRowIndex;
        //        datarange.StartColumnIndex = PivotTable.Location.Reference.StartColumnIndex;
        //        datarange.EndRowIndex = PivotTable.Location.Reference.EndRowIndex;
        //        datarange.EndColumnIndex = PivotTable.Location.Reference.EndColumnIndex;
        //    }

        //    SLCacheField cfTemp;

        //    for (j = datarange.StartColumnIndex; j <= datarange.EndColumnIndex; ++j)
        //    {
        //        pt = new SLCellPoint(datarange.StartRowIndex, j);
        //        if (cells.ContainsKey(pt))
        //        {
        //            cfTemp = new SLCacheField();
        //            cell = cells[pt];
        //            if (cell.DataType == CellValues.SharedString)
        //            {
        //                if (cell.CellText != null) iSharedStringIndex = Convert.ToInt32(cell.CellText);
        //                else iSharedStringIndex = Convert.ToInt32(cell.NumericValue);

        //                if (dictSharedStringSimplified.ContainsKey(iSharedStringIndex))
        //                {
        //                    cfTemp.Name = dictSharedStringSimplified[iSharedStringIndex];
        //                }
        //                else
        //                {
        //                    rst.FromHash(listSharedString[iSharedStringIndex]);
        //                    sPlainString = rst.ToPlainString();
        //                    dictSharedStringSimplified[iSharedStringIndex] = sPlainString;
        //                    cfTemp.Name = sPlainString;
        //                }
        //            }
        //            else if (cell.DataType == CellValues.Number)
        //            {
        //            }
        //            else if (cell.DataType == CellValues.Boolean)
        //            {
        //            }
        //            else
        //            {
        //            }
        //        }
        //        else
        //        {
        //        }
        //    }

        //    cachedef.PivotCacheDefinition = pcd.ToPivotCacheDefinition();
        //    #endregion Content for PivotCacheDefinition

        //    PivotTableCacheRecordsPart cacherecords = cachedef.AddNewPart<PivotTableCacheRecordsPart>();

        //    #region Content for PivotCacheRecords
        //    #endregion Content for PivotCacheRecords

        //    //TODO check string.isNull for relID
        //    WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
        //    PivotTablePart ptp = wsp.AddNewPart<PivotTablePart>();

        //    #region Content for PivotTableDefinition
        //    #endregion Content for PivotTableDefinition

        //    ptp.AddPart(cachedef);

        //    if (!SelectNewWorksheet)
        //    {
        //        SelectWorksheet(sCurrentWorksheetName);
        //    }

        //    return false;
        //}
    }
}
