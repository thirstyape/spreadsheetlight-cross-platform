using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using SLA = SpreadsheetLight.Drawing;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// Adds a new worksheet, and selects the new worksheet as the active one.
        /// </summary>
        /// <param name="WorksheetName">The name should not be blank, nor exceed 31 characters. And it cannot contain these characters: \/?*[] It cannot be the same as an existing name (case-insensitive). But there's nothing stopping you from using 3 spaces as a name.</param>
        /// <returns>True if the name is valid and the worksheet is successfully added. False otherwise.</returns>
        public bool AddWorksheet(string WorksheetName)
        {
            if (!SLTool.CheckSheetChartName(WorksheetName))
            {
                return false;
            }
            foreach (SLSheet sheet in slwb.Sheets)
            {
                if (sheet.Name.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            // if there's at least one worksheet, then there's at least a selected worksheet.
            if (slwb.Sheets.Count > 0)
            {
                WriteSelectedWorksheet();
            }

            gsSelectedWorksheetName = WorksheetName;
            ++giWorksheetIdCounter;

            // use an empty string for the Id first
            slwb.Sheets.Add(new SLSheet(WorksheetName, (uint)giWorksheetIdCounter, string.Empty, SLSheetType.Worksheet));

            giSelectedWorksheetID = (uint)giWorksheetIdCounter;

            gsSelectedWorksheetRelationshipID = string.Empty;

            slws = new SLWorksheet(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors, SimpleTheme.ThemeColumnWidth, SimpleTheme.ThemeColumnWidthInEMU, SimpleTheme.ThemeMaxDigitWidth, SimpleTheme.listColumnStepSize, SimpleTheme.ThemeRowHeight);
            IsNewWorksheet = true;

            return true;
        }

        /// <summary>
        /// Select an existing worksheet. If the given name doesn't match an existing worksheet, the previously selected worksheet is used.
        /// </summary>
        /// <param name="WorksheetName">The name of an existing worksheet.</param>
        /// <returns>True if there's an existing worksheet with that name and that worksheet is successfully selected. False otherwise.</returns>
        public bool SelectWorksheet(string WorksheetName)
        {
            // if the current worksheet is already selected, no need to select again, right?
            if (gsSelectedWorksheetName.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
                // originally return false. A fellow developer contacted me that this doesn't quite
                // fit the explanation of the return value. And I agree.
                // Frankly, I didn't expect anyone to actually use the return value...
                // I should've used a void as the return type, but now it's too late...
            }

            uint iNewlySelectedWorksheetID = 0;
            string sNewlySelectedWorksheetRelationshipID = string.Empty;
            bool bFound = false;
            for (int i = 0; i < slwb.Sheets.Count; ++i)
            {
                if (slwb.Sheets[i].Name.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase)
                    && slwb.Sheets[i].SheetType == SLSheetType.Worksheet)
                {
                    bFound = true;
                    iNewlySelectedWorksheetID = slwb.Sheets[i].SheetId;
                    sNewlySelectedWorksheetRelationshipID = slwb.Sheets[i].Id;
                    break;
                }
            }
            if (!bFound)
            {
                return false;
            }

            // if there's at least one worksheet, then there's at least a selected worksheet.
            if (slwb.Sheets.Count > 0)
            {
                WriteSelectedWorksheet();
            }

            giSelectedWorksheetID = iNewlySelectedWorksheetID;
            gsSelectedWorksheetName = WorksheetName;
            gsSelectedWorksheetRelationshipID = sNewlySelectedWorksheetRelationshipID;

            LoadSelectedWorksheet();
            IsNewWorksheet = false;

            return true;
        }

        /// <summary>
        /// Show the worksheet (a.k.a unhide worksheet). This includes chart sheets, dialog sheets and macro sheets.
        /// </summary>
        /// <param name="WorksheetName">The name of the worksheet.</param>
        public void ShowWorksheet(string WorksheetName)
        {
            this.ShowHideWorksheet(WorksheetName, SheetStateValues.Visible);
        }

        /// <summary>
        /// Hide the worksheet. This includes chart sheets, dialog sheets and macro sheets.
        /// </summary>
        /// <param name="WorksheetName">The name of the worksheet.</param>
        public void HideWorksheet(string WorksheetName)
        {
            this.ShowHideWorksheet(WorksheetName, SheetStateValues.Hidden);
        }

        /// <summary>
        /// Hide the worksheet. This includes chart sheets, dialog sheets and macro sheets.
        /// </summary>
        /// <param name="WorksheetName">The name of the worksheet.</param>
        /// <param name="IsVeryHidden">True to set the worksheet as very hidden. False otherwise.</param>
        public void HideWorksheet(string WorksheetName, bool IsVeryHidden)
        {
            this.ShowHideWorksheet(WorksheetName, IsVeryHidden ? SheetStateValues.VeryHidden : SheetStateValues.Hidden);
        }

        private void ShowHideWorksheet(string WorksheetName, SheetStateValues State)
        {
            for (int i = 0; i < slwb.Sheets.Count; ++i)
            {
                if (slwb.Sheets[i].Name.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    slwb.Sheets[i].State = State;
                    break;
                }
            }
        }

        /// <summary>
        /// Indicates if the worksheet is hidden. Note that if the worksheet name isn't of an existing worksheet, the return value is true. Think of a non-existent worksheet as very very hidden, like hidden in another dimension or something.
        /// </summary>
        /// <param name="WorksheetName">The name of the worksheet.</param>
        /// <returns>True if the worksheet is hidden. False otherwise.</returns>
        public bool IsWorksheetHidden(string WorksheetName)
        {
            bool result = true;
            for (int i = 0; i < slwb.Sheets.Count; ++i)
            {
                if (slwb.Sheets[i].Name.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase)
                    && slwb.Sheets[i].State == SheetStateValues.Visible)
                {
                    result = false;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Rename an existing worksheet. This includes chart sheets, dialog sheets and macro sheets.
        /// </summary>
        /// <param name="ExistingWorksheetName">The name of the existing worksheet.</param>
        /// <param name="NewWorksheetName">The new name for the existing worksheet. The name should not be blank, nor exceed 31 characters. And it cannot contain these characters: \/?*[] It cannot be the same as an existing name (case-insensitive).</param>
        /// <returns>True if renaming is successful. False otherwise.</returns>
        public bool RenameWorksheet(string ExistingWorksheetName, string NewWorksheetName)
        {
            if (!SLTool.CheckSheetChartName(NewWorksheetName))
            {
                return false;
            }
            foreach (SLSheet sheet in slwb.Sheets)
            {
                if (sheet.Name.Equals(NewWorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            int i;
            bool bFound = false;
            for (i = 0; i < slwb.Sheets.Count; ++i)
            {
                if (slwb.Sheets[i].Name.Equals(ExistingWorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    slwb.Sheets[i].Name = NewWorksheetName;
                    bFound = true;
                    break;
                }
            }
            if (!bFound)
            {
                return false;
            }

            if (gsSelectedWorksheetName.Equals(ExistingWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                gsSelectedWorksheetName = NewWorksheetName;
            }

            // External reference
            // =SUM([Budget.xlsx]Annual!C10:C25)
            // When the source is not open, the external reference includes the entire path.
            // External reference
            // =SUM('C:\Reports\[Budget.xlsx]Annual'!C10:C25)

            // need to check for invariant/case-insensitive?
            // I'm going to partially ignore workbook references, meaning [Budget.xlsx]Annual! is still
            // kind of ok..
            // This is because if the sheet name is [Budget.xlsx]1999's Annual then it becomes
            // "'[Budget.xlsx]1999''s Annual'". At least I think that's what it becomes.
            // The rules for single quotes are ... obscure...
            // Hopefully there aren't any workbook references at all.
            string sPattern = @"(?<pre>^|[^\]])" + SLTool.FormatWorksheetNameForFormula(ExistingWorksheetName) + "!";
            string sReplacement = "${pre}" + SLTool.FormatWorksheetNameForFormula(NewWorksheetName) + "!";
            foreach (SLDefinedName dn in slwb.DefinedNames)
            {
                dn.Text = Regex.Replace(dn.Text, sPattern, sReplacement);
            }

            // TODO update for cells in other worksheets
            // One of the few cases where having a shared formula repository is useful... (like shared strings)
            SLCell c;
            List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
            List<int> listColumnKeys;
            foreach (int rowkey in listRowKeys)
            {
                listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                foreach (int colkey in listColumnKeys)
                {
                    c = slws.CellWarehouse.Cells[rowkey][colkey];
                    if (c.CellText != null && c.CellText.StartsWith("="))
                    {
                        c.CellText = Regex.Replace(c.CellText, sPattern, sReplacement);
                    }
                    if (c.CellFormula != null)
                    {
                        c.CellFormula.FormulaText = Regex.Replace(c.CellFormula.FormulaText, sPattern, sReplacement);
                    }
                    slws.CellWarehouse.SetValue(rowkey, colkey, c);
                }
            }

            // this updates an chart references
            foreach (WorksheetPart wsp in xl.WorkbookPart.WorksheetParts)
            {
                if (wsp.DrawingsPart != null)
                {
                    C.Chart chartOriginal;
                    C.Chart chartNew;
                    // I'm looking for something like the following:
                    //<c:f>AwesomeName!$A$2:$A$4</c:f>
                    // An alternative might be:
                    //<f>AwesomeName!$A$2:$A$4</f>
                    // In case the XML tag prefix is not there, hence the smiley face <:
                    // I'm replacing the entire inner XML, so it's best to be as specific as possible.
                    string sReplaceExisting = "(?<pre>[<:]f>)" + SLTool.FormatWorksheetNameForFormula(ExistingWorksheetName) + "(?<post>!\\$[A-Z]{1,3}\\$)";
                    string sReplaceNew = "${pre}" + SLTool.FormatWorksheetNameForFormula(NewWorksheetName) + "${post}";
                    foreach (ChartPart cp in wsp.DrawingsPart.ChartParts)
                    {
                        chartOriginal = cp.ChartSpace.Elements<C.Chart>().First();
                        chartNew = (C.Chart)chartOriginal.CloneNode(true);
                        if (Regex.IsMatch(chartNew.InnerXml, sReplaceExisting))
                        {
                            chartNew.InnerXml = Regex.Replace(chartNew.InnerXml, sReplaceExisting, sReplaceNew);
                            cp.ChartSpace.ReplaceChild<C.Chart>(chartNew, chartOriginal);
                            cp.ChartSpace.Save();
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Copy the contents of a worksheet to another worksheet. The current worksheet cannot be the source of the copy operation, nor the destination of the copy operation.
        /// </summary>
        /// <param name="ExistingWorksheetName">The worksheet to be copied from. This cannot be the current worksheet.</param>
        /// <param name="NewWorksheetName">The worksheet to be copied to. If this doesn't exist, a new worksheet is created. If it's an existing worksheet, the contents of the existing worksheet will be overwritten. The new worksheet cannot be the currently selected worksheet.</param>
        /// <returns>True if copying is successful. False otherwise</returns>
        public bool CopyWorksheet(string ExistingWorksheetName, string NewWorksheetName)
        {
            if (ExistingWorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase)
                || NewWorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase)
                || ExistingWorksheetName.Equals(NewWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            bool bExistingFound = false;
            string sExistingRelId = string.Empty;
            bool bNewFound = false;
            string sNewRelId = string.Empty;
            foreach (SLSheet sheet in slwb.Sheets)
            {
                if (sheet.Name.Equals(ExistingWorksheetName, StringComparison.OrdinalIgnoreCase)
                    && sheet.SheetType == SLSheetType.Worksheet)
                {
                    sExistingRelId = sheet.Id;
                    bExistingFound = true;
                }
                else if (sheet.Name.Equals(NewWorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    sNewRelId = sheet.Id;
                    bNewFound = true;
                }
            }

            // no existing worksheet found! So just return. Nothing to copy!
            if (!bExistingFound) return false;

            // destination worksheet not an existing worksheet, but the name failed the test
            if (!bNewFound && !SLTool.CheckSheetChartName(NewWorksheetName)) return false;

            WorksheetPart wspExisting = (WorksheetPart)wbp.GetPartById(sExistingRelId);
            WorksheetPart wspNew = wbp.AddNewPart<WorksheetPart>();

            using (StreamReader sr = new StreamReader(wspExisting.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(wspNew.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            ImagePart imgpNew;

            #region ControlPropertiesParts
            if (wspExisting.ControlPropertiesParts != null)
            {
                ControlPropertiesPart cppNew;
                foreach (ControlPropertiesPart cpp in wspExisting.ControlPropertiesParts)
                {
                    cppNew = wspNew.AddNewPart<ControlPropertiesPart>(wspExisting.GetIdOfPart(cpp));
                    using (StreamReader sr = new StreamReader(cpp.GetStream()))
                    {
                        using (StreamWriter sw = new StreamWriter(cppNew.GetStream(FileMode.Create)))
                        {
                            sw.Write(sr.ReadToEnd());
                        }
                    }
                }
            }
            #endregion

            #region CustomPropertyParts
            if (wspExisting.CustomPropertyParts != null)
            {
                CustomPropertyPart cppNew;
                foreach (CustomPropertyPart cpp in wspExisting.CustomPropertyParts)
                {
                    cppNew = wspNew.AddCustomPropertyPart(cpp.ContentType, wspExisting.GetIdOfPart(cpp));
                    using (StreamReader sr = new StreamReader(cpp.GetStream()))
                    {
                        using (StreamWriter sw = new StreamWriter(cppNew.GetStream(FileMode.Create)))
                        {
                            sw.Write(sr.ReadToEnd());
                        }
                    }
                }
            }
            #endregion

            #region DrawingsPart
            if (wspExisting.DrawingsPart != null)
            {
                DrawingsPart dp = wspExisting.DrawingsPart;
                DrawingsPart dpNew = wspNew.AddNewPart<DrawingsPart>(wspExisting.GetIdOfPart(dp));
                using (StreamReader sr = new StreamReader(dp.GetStream()))
                {
                    using (StreamWriter sw = new StreamWriter(dpNew.GetStream(FileMode.Create)))
                    {
                        sw.Write(sr.ReadToEnd());
                    }
                }

                #region DrawingsPart ChartParts
                ChartPart cpNew;
                foreach (ChartPart cp in dp.ChartParts)
                {
                    cpNew = dpNew.AddNewPart<ChartPart>(dp.GetIdOfPart(cp));
                    this.FeedDataChartPart(cpNew, cp);
                }
                #endregion

                #region DrawingsPart CustomXmlParts
                if (dp.CustomXmlParts != null)
                {
                    CustomXmlPart cxpNew;
                    foreach (CustomXmlPart cxp in dp.CustomXmlParts)
                    {
                        cxpNew = dpNew.AddCustomXmlPart(cxp.ContentType, dp.GetIdOfPart(cxp));
                        this.FeedDataCustomXmlPart(cxpNew, cxp);
                    }
                }
                #endregion

                #region DrawingsPart DiagramColorsParts
                if (dp.DiagramColorsParts != null)
                {
                    DiagramColorsPart dcpNew;
                    foreach (DiagramColorsPart dcp in dp.DiagramColorsParts)
                    {
                        dcpNew = dpNew.AddNewPart<DiagramColorsPart>(dp.GetIdOfPart(dcp));
                        this.FeedDataDiagramColorsPart(dcpNew, dcp);
                    }
                }
                #endregion

                #region DrawingsPart DiagramDataParts
                if (dp.DiagramDataParts != null)
                {
                    DiagramDataPart ddpNew;
                    foreach (DiagramDataPart ddp in dp.DiagramDataParts)
                    {
                        ddpNew = dpNew.AddNewPart<DiagramDataPart>(dp.GetIdOfPart(ddp));
                        this.FeedDataDiagramDataPart(ddpNew, ddp);
                    }
                }
                #endregion

                #region DrawingsPart DiagramLayoutDefinitionParts
                if (dp.DiagramLayoutDefinitionParts != null)
                {
                    DiagramLayoutDefinitionPart dldpNew;
                    foreach (DiagramLayoutDefinitionPart dldp in dp.DiagramLayoutDefinitionParts)
                    {
                        dldpNew = dpNew.AddNewPart<DiagramLayoutDefinitionPart>(dp.GetIdOfPart(dldp));
                        this.FeedDataDiagramLayoutDefinitionPart(dldpNew, dldp);
                    }
                }
                #endregion

                #region DrawingsPart DiagramPersistLayoutParts
                if (dp.DiagramPersistLayoutParts != null)
                {
                    DiagramPersistLayoutPart dplpNew;
                    foreach (DiagramPersistLayoutPart dplp in dp.DiagramPersistLayoutParts)
                    {
                        dplpNew = dpNew.AddNewPart<DiagramPersistLayoutPart>(dp.GetIdOfPart(dplp));
                        this.FeedDataDiagramPersistLayoutPart(dplpNew, dplp);
                    }
                }
                #endregion

                #region DrawingsPart DiagramStyleParts
                if (dp.DiagramStyleParts != null)
                {
                    DiagramStylePart dspNew;
                    foreach (DiagramStylePart dsp in dp.DiagramStyleParts)
                    {
                        dspNew = dpNew.AddNewPart<DiagramStylePart>(dp.GetIdOfPart(dsp));
                        this.FeedDataDiagramStylePart(dspNew, dsp);
                    }
                }
                #endregion

                #region DrawingsPart ImageParts
                foreach (ImagePart imgp in dp.ImageParts)
                {
                    imgpNew = dpNew.AddImagePart(imgp.ContentType, dp.GetIdOfPart(imgp));
                    this.FeedDataImagePart(imgpNew, imgp);
                }
                #endregion
            }
            #endregion

            #region EmbeddedControlPersistenceBinaryDataParts
            if (wspExisting.EmbeddedControlPersistenceBinaryDataParts != null)
            {
                EmbeddedControlPersistenceBinaryDataPart binNew;
                foreach (EmbeddedControlPersistenceBinaryDataPart bin in wspExisting.EmbeddedControlPersistenceBinaryDataParts)
                {
                    binNew = wspNew.AddEmbeddedControlPersistenceBinaryDataPart(bin.ContentType, wspExisting.GetIdOfPart(bin));
                    this.FeedDataEmbeddedControlPersistenceBinaryDataPart(binNew, bin);
                }
            }
            #endregion

            #region EmbeddedControlPersistenceParts
            if (wspExisting.EmbeddedControlPersistenceParts != null)
            {
                EmbeddedControlPersistencePart ecppNew;
                foreach (EmbeddedControlPersistencePart ecpp in wspExisting.EmbeddedControlPersistenceParts)
                {
                    ecppNew = wspNew.AddEmbeddedControlPersistencePart(ecpp.ContentType, wspExisting.GetIdOfPart(ecpp));
                    this.FeedDataEmbeddedControlPersistencePart(ecppNew, ecpp);
                }
            }
            #endregion

            #region EmbeddedObjectParts
            if (wspExisting.EmbeddedObjectParts != null)
            {
                EmbeddedObjectPart eopNew;
                foreach (EmbeddedObjectPart eop in wspExisting.EmbeddedObjectParts)
                {
                    eopNew = wspNew.AddEmbeddedObjectPart(eop.ContentType);
                    this.FeedDataEmbeddedObjectPart(eopNew, eop);
                }
            }
            #endregion

            #region EmbeddedPackageParts
            if (wspExisting.EmbeddedPackageParts != null)
            {
                EmbeddedPackagePart eppNew;
                foreach (EmbeddedPackagePart epp in wspExisting.EmbeddedPackageParts)
                {
                    eppNew = wspNew.AddEmbeddedPackagePart(epp.ContentType);
                    this.FeedDataEmbeddedPackagePart(eppNew, epp);
                }
            }
            #endregion

            #region HyperlinkRelationships
            if (wspExisting.HyperlinkRelationships.Count() > 0)
            {
                foreach (HyperlinkRelationship hlrel in wspExisting.HyperlinkRelationships)
                {
                    wspNew.AddHyperlinkRelationship(hlrel.Uri, hlrel.IsExternal, hlrel.Id);
                }
            }
            #endregion

            #region ImageParts
            if (wspExisting.ImageParts != null)
            {
                foreach (ImagePart imgp in wspExisting.ImageParts)
                {
                    imgpNew = wspNew.AddImagePart(imgp.ContentType, wspExisting.GetIdOfPart(imgp));
                    this.FeedDataImagePart(imgpNew, imgp);
                }
            }
            #endregion

            #region PivotTableParts
            if (wspExisting.PivotTableParts != null)
            {
                PivotTablePart ptpNew;
                foreach (PivotTablePart ptp in wspExisting.PivotTableParts)
                {
                    ptpNew = wspNew.AddNewPart<PivotTablePart>(wspExisting.GetIdOfPart(ptp));
                    using (StreamReader sr = new StreamReader(ptp.GetStream()))
                    {
                        using (StreamWriter sw = new StreamWriter(ptpNew.GetStream(FileMode.Create)))
                        {
                            sw.Write(sr.ReadToEnd());
                        }
                    }

                    if (ptp.PivotTableCacheDefinitionPart != null)
                    {
                        ptpNew.AddNewPart<PivotTableCacheDefinitionPart>(ptp.GetIdOfPart(ptp.PivotTableCacheDefinitionPart));
                        using (StreamReader sr = new StreamReader(ptp.PivotTableCacheDefinitionPart.GetStream()))
                        {
                            using (StreamWriter sw = new StreamWriter(ptpNew.PivotTableCacheDefinitionPart.GetStream(FileMode.Create)))
                            {
                                sw.Write(sr.ReadToEnd());
                            }
                        }
                    }
                }
            }
            #endregion

            #region QueryTableParts
            if (wspExisting.QueryTableParts != null)
            {
                QueryTablePart qtpNew;
                foreach (QueryTablePart qtp in wspExisting.QueryTableParts)
                {
                    qtpNew = wspNew.AddNewPart<QueryTablePart>(wspExisting.GetIdOfPart(qtp));
                    using (StreamReader sr = new StreamReader(qtp.GetStream()))
                    {
                        using (StreamWriter sw = new StreamWriter(qtpNew.GetStream(FileMode.Create)))
                        {
                            sw.Write(sr.ReadToEnd());
                        }
                    }
                }
            }
            #endregion

            #region SingleCellTablePart
            if (wspExisting.SingleCellTablePart != null)
            {
                wspNew.AddNewPart<SingleCellTablePart>(wspExisting.GetIdOfPart(wspExisting.SingleCellTablePart));
                using (StreamReader sr = new StreamReader(wspExisting.SingleCellTablePart.GetStream()))
                {
                    using (StreamWriter sw = new StreamWriter(wspNew.SingleCellTablePart.GetStream(FileMode.Create)))
                    {
                        sw.Write(sr.ReadToEnd());
                    }
                }
            }
            #endregion

            #region SlicersParts
            if (wspExisting.SlicersParts != null)
            {
                SlicersPart spNew;
                foreach (SlicersPart sp in wspExisting.SlicersParts)
                {
                    spNew = wspNew.AddNewPart<SlicersPart>(wspExisting.GetIdOfPart(sp));
                    using (StreamReader sr = new StreamReader(sp.GetStream()))
                    {
                        using (StreamWriter sw = new StreamWriter(spNew.GetStream(FileMode.Create)))
                        {
                            sw.Write(sr.ReadToEnd());
                        }
                    }
                }
            }
            #endregion

            #region SpreadsheetPrinterSettingsParts
            // What is inside this thing?!? An error occurs in the copied worksheet if the
            // existing worksheet had tables. The copied worksheet doesn't need the printer
            // settings of the original worksheet, right?
            //if (wspExisting.SpreadsheetPrinterSettingsParts != null)
            //{
            //    SpreadsheetPrinterSettingsPart spspNew;
            //    foreach (SpreadsheetPrinterSettingsPart spsp in wspExisting.SpreadsheetPrinterSettingsParts)
            //    {
            //        spspNew = wspNew.AddNewPart<SpreadsheetPrinterSettingsPart>(wspExisting.GetIdOfPart(spsp));
            //        using (StreamReader sr = new StreamReader(spsp.GetStream()))
            //        {
            //            using (StreamWriter sw = new StreamWriter(spspNew.GetStream(FileMode.Create)))
            //            {
            //                sw.Write(sr.ReadToEnd());
            //            }
            //        }
            //    }
            //}
            #endregion

            #region TableDefinitionParts
            if (wspExisting.TableDefinitionParts != null)
            {
                TableDefinitionPart tdpNew;
                QueryTablePart qtpNew;
                foreach (TableDefinitionPart tdp in wspExisting.TableDefinitionParts)
                {
                    tdpNew = wspNew.AddNewPart<TableDefinitionPart>(wspExisting.GetIdOfPart(tdp));
                    using (StreamReader sr = new StreamReader(tdp.GetStream()))
                    {
                        using (StreamWriter sw = new StreamWriter(tdpNew.GetStream(FileMode.Create)))
                        {
                            sw.Write(sr.ReadToEnd());
                        }
                    }
                    slwb.RefreshPossibleTableId();
                    tdpNew.Table.Id = slwb.PossibleTableId;
                    tdpNew.Table.DisplayName = slwb.GetNextPossibleTableName();
                    tdpNew.Table.Name = tdpNew.Table.DisplayName.Value;

                    foreach (QueryTablePart qtp in tdp.QueryTableParts)
                    {
                        qtpNew = tdpNew.AddNewPart<QueryTablePart>(tdp.GetIdOfPart(qtp));
                        using (StreamReader sr = new StreamReader(qtp.GetStream()))
                        {
                            using (StreamWriter sw = new StreamWriter(qtpNew.GetStream(FileMode.Create)))
                            {
                                sw.Write(sr.ReadToEnd());
                            }
                        }
                    }
                }
            }
            #endregion

            #region VmlDrawingParts
            if (wspExisting.VmlDrawingParts != null)
            {
                VmlDrawingPart vdpNew;
                foreach (VmlDrawingPart vdp in wspExisting.VmlDrawingParts)
                {
                    vdpNew = wspNew.AddNewPart<VmlDrawingPart>(wspExisting.GetIdOfPart(vdp));
                    this.FeedDataVmlDrawingPart(vdpNew, vdp);
                }
            }
            #endregion

            #region WorksheetCommentsPart
            if (wspExisting.WorksheetCommentsPart != null)
            {
                wspNew.AddNewPart<WorksheetCommentsPart>(wspExisting.GetIdOfPart(wspExisting.WorksheetCommentsPart));
                using (StreamReader sr = new StreamReader(wspExisting.WorksheetCommentsPart.GetStream()))
                {
                    using (StreamWriter sw = new StreamWriter(wspNew.WorksheetCommentsPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(sr.ReadToEnd());
                    }
                }
            }
            #endregion

            #region WorksheetSortMapPart
            if (wspExisting.WorksheetSortMapPart != null)
            {
                wspNew.AddNewPart<WorksheetSortMapPart>(wspExisting.GetIdOfPart(wspExisting.WorksheetSortMapPart));
                using (StreamReader sr = new StreamReader(wspExisting.WorksheetSortMapPart.GetStream()))
                {
                    using (StreamWriter sw = new StreamWriter(wspNew.WorksheetSortMapPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(sr.ReadToEnd());
                    }
                }
            }
            #endregion

            if (!bNewFound)
            {
                ++giWorksheetIdCounter;
                slwb.Sheets.Add(new SLSheet(NewWorksheetName, (uint)giWorksheetIdCounter, wbp.GetIdOfPart(wspNew), SLSheetType.Worksheet));
            }
            else
            {
                wbp.DeletePart(sNewRelId);

                foreach (SLSheet sheet in slwb.Sheets)
                {
                    if (sheet.Name.Equals(NewWorksheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.Id = wbp.GetIdOfPart(wspNew);
                    }
                }
            }

            // 5 Mar 2017: bug fix, return value used to be always false... bugger...
            return true;
        }

        /// <summary>
        /// Delete a worksheet. The currently selected worksheet cannot be deleted.
        /// </summary>
        /// <param name="WorksheetName">The name of the worksheet to be deleted.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool DeleteWorksheet(string WorksheetName)
        {
            bool result = false;
            if (!WorksheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                string sRelId = string.Empty;
                int i = 0;
                for (i = 0; i < slwb.Sheets.Count; ++i)
                {
                    if (slwb.Sheets[i].Name.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        sRelId = slwb.Sheets[i].Id;
                        break;
                    }
                }

                if (sRelId.Length > 0)
                {
                    result = true;

                    wbp.DeletePart(sRelId);
                    slwb.Sheets.RemoveAt(i);
                }
            }

            return result;
        }

        /// <summary>
        /// Get the worksheet name of the currently selected worksheet.
        /// </summary>
        /// <returns>The currently selected worksheet name.</returns>
        public string GetCurrentWorksheetName()
        {
            return gsSelectedWorksheetName;
        }

        /// <summary>
        /// Get a list of names of existing worksheets currently in the spreadsheet, excluding chart sheets, macro sheets and dialog sheets.
        /// This is identical to GetSheetNames() but is more appropriately named.
        /// </summary>
        /// <returns>A list of names of existing worksheets.</returns>
        public List<string> GetWorksheetNames()
        {
            return this.GetSheetNames(false);
        }

        /// <summary>
        /// Get a list of names of existing worksheets currently in the spreadsheet, excluding chart sheets, macro sheets and dialog sheets.
        /// This is identical to GetWorksheetNames() but is kept to maintain function overload consistency with the IncludeAll parameter version.
        /// </summary>
        /// <returns>A list of names of existing worksheets.</returns>
        public List<string> GetSheetNames()
        {
            return this.GetSheetNames(false);
        }

        /// <summary>
        /// Get a list of names of existing sheets currently in the spreadsheet.
        /// </summary>
        /// <param name="IncludeAll">True to include chart sheets, macro sheets and dialog sheets. False to limit to only worksheets.</param>
        /// <returns>A list of names of existing sheets.</returns>
        public List<string> GetSheetNames(bool IncludeAll)
        {
            List<string> list = new List<string>();

            if (IncludeAll)
            {
                for (int i = 0; i < slwb.Sheets.Count; ++i)
                {
                    list.Add(slwb.Sheets[i].Name);
                }
            }
            else
            {
                // worksheets only
                for (int i = 0; i < slwb.Sheets.Count; ++i)
                {
                    if (slwb.Sheets[i].SheetType == SLSheetType.Worksheet) list.Add(slwb.Sheets[i].Name);
                }
            }

            return list;
        }

        /// <summary>
        /// Get statistical information on the currently selected worksheet. NOTE: The information is only current at point of retrieval.
        /// </summary>
        /// <returns>An SLWorksheetStatistics object with the information.</returns>
        public SLWorksheetStatistics GetWorksheetStatistics()
        {
            SLWorksheetStatistics wsstats = new SLWorksheetStatistics();

            if (slws.CellWarehouse.Cells.Count > 0)
            {
                HashSet<int> hsRows = new HashSet<int>();
                HashSet<int> hsColumns = new HashSet<int>();

                List<int> listRowKeys = slws.CellWarehouse.Cells.Keys.ToList<int>();
                List<int> listColumnKeys;

                listRowKeys.Sort();

                if (listRowKeys.Count > 0)
                {
                    wsstats.iStartRowIndex = listRowKeys[0];
                    wsstats.iEndRowIndex = listRowKeys[0];

                    listColumnKeys = slws.CellWarehouse.Cells[listRowKeys[0]].Keys.ToList<int>();

                    if (listColumnKeys.Count > 0)
                    {
                        wsstats.iStartColumnIndex = listColumnKeys[0];
                        wsstats.iEndColumnIndex = listColumnKeys[0];
                    }
                }

                wsstats.iNumberOfCells = 0;
                wsstats.iNumberOfEmptyCells = 0;

                SLCell c;
                foreach (int rowkey in listRowKeys)
                {
                    listColumnKeys = slws.CellWarehouse.Cells[rowkey].Keys.ToList<int>();
                    foreach (int colkey in listColumnKeys)
                    {
                        if (rowkey < wsstats.iStartRowIndex) wsstats.iStartRowIndex = rowkey;
                        if (rowkey > wsstats.iEndRowIndex) wsstats.iEndRowIndex = rowkey;
                        if (colkey < wsstats.iStartColumnIndex) wsstats.iStartColumnIndex = colkey;
                        if (colkey > wsstats.iEndColumnIndex) wsstats.iEndColumnIndex = colkey;

                        if (!hsRows.Contains(rowkey)) hsRows.Add(rowkey);
                        if (!hsColumns.Contains(colkey)) hsColumns.Add(colkey);

                        ++wsstats.iNumberOfCells;

                        c = slws.CellWarehouse.Cells[rowkey][colkey];
                        if (c.CellText != null && c.CellText.Length == 0 && c.CellFormula == null)
                        {
                            ++wsstats.iNumberOfEmptyCells;
                        }
                    }
                }

                foreach (int rowindex in slws.RowProperties.Keys)
                {
                    if (rowindex < wsstats.iStartRowIndex) wsstats.iStartRowIndex = rowindex;
                    if (rowindex > wsstats.iEndRowIndex) wsstats.iEndRowIndex = rowindex;
                    if (!hsRows.Contains(rowindex)) hsRows.Add(rowindex);
                }

                foreach (int columnindex in slws.ColumnProperties.Keys)
                {
                    if (columnindex < wsstats.iStartColumnIndex) wsstats.iStartColumnIndex = columnindex;
                    if (columnindex > wsstats.iEndColumnIndex) wsstats.iEndColumnIndex = columnindex;
                    if (!hsColumns.Contains(columnindex)) hsColumns.Add(columnindex);
                }

                wsstats.iNumberOfRows = hsRows.Count;
                wsstats.iNumberOfColumns = hsColumns.Count;
            }

            return wsstats;
        }

        internal void FeedDataImagePart(ImagePart NewPart, ImagePart ExistingPart)
        {
            System.Drawing.Imaging.ImageFormat imgtype = SLTool.TranslateImageContentType(ExistingPart.ContentType);
            using (System.Drawing.Bitmap bm = new System.Drawing.Bitmap(ExistingPart.GetStream()))
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    bm.Save(ms, imgtype);
                    ms.Position = 0;
                    NewPart.FeedData(ms);
                }
            }
        }

        internal void FeedDataSlidePart(SlidePart NewPart, SlidePart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            ImagePart imgpNew;

            if (ExistingPart.ChartParts != null)
            {
                ChartPart cpNew;
                foreach (ChartPart cp in ExistingPart.ChartParts)
                {
                    cpNew = NewPart.AddNewPart<ChartPart>(ExistingPart.GetIdOfPart(cp));
                    this.FeedDataChartPart(cpNew, cp);
                }
            }

            if (ExistingPart.CustomXmlParts != null)
            {
                CustomXmlPart cxpNew;
                foreach (CustomXmlPart cxp in ExistingPart.CustomXmlParts)
                {
                    cxpNew = NewPart.AddCustomXmlPart(cxp.ContentType, ExistingPart.GetIdOfPart(cxp));
                    this.FeedDataCustomXmlPart(cxpNew, cxp);
                }
            }

            if (ExistingPart.DiagramColorsParts != null)
            {
                DiagramColorsPart dcpNew;
                foreach (DiagramColorsPart dcp in ExistingPart.DiagramColorsParts)
                {
                    dcpNew = NewPart.AddNewPart<DiagramColorsPart>(ExistingPart.GetIdOfPart(dcp));
                    this.FeedDataDiagramColorsPart(dcpNew, dcp);
                }
            }

            if (ExistingPart.DiagramDataParts != null)
            {
                DiagramDataPart ddpNew;
                foreach (DiagramDataPart ddp in ExistingPart.DiagramDataParts)
                {
                    ddpNew = NewPart.AddNewPart<DiagramDataPart>(ExistingPart.GetIdOfPart(ddp));
                    this.FeedDataDiagramDataPart(ddpNew, ddp);
                }
            }

            if (ExistingPart.DiagramLayoutDefinitionParts != null)
            {
                DiagramLayoutDefinitionPart dldpNew;
                foreach (DiagramLayoutDefinitionPart dldp in ExistingPart.DiagramLayoutDefinitionParts)
                {
                    dldpNew = NewPart.AddNewPart<DiagramLayoutDefinitionPart>(ExistingPart.GetIdOfPart(dldp));
                    this.FeedDataDiagramLayoutDefinitionPart(dldpNew, dldp);
                }
            }

            if (ExistingPart.DiagramPersistLayoutParts != null)
            {
                DiagramPersistLayoutPart dplpNew;
                foreach (DiagramPersistLayoutPart dplp in ExistingPart.DiagramPersistLayoutParts)
                {
                    dplpNew = NewPart.AddNewPart<DiagramPersistLayoutPart>(ExistingPart.GetIdOfPart(dplp));
                    this.FeedDataDiagramPersistLayoutPart(dplpNew, dplp);
                }
            }

            if (ExistingPart.DiagramStyleParts != null)
            {
                DiagramStylePart dspNew;
                foreach (DiagramStylePart dsp in ExistingPart.DiagramStyleParts)
                {
                    dspNew = NewPart.AddNewPart<DiagramStylePart>(ExistingPart.GetIdOfPart(dsp));
                    this.FeedDataDiagramStylePart(dspNew, dsp);
                }
            }

            if (ExistingPart.EmbeddedControlPersistenceBinaryDataParts != null)
            {
                EmbeddedControlPersistenceBinaryDataPart binNew;
                foreach (EmbeddedControlPersistenceBinaryDataPart bin in ExistingPart.EmbeddedControlPersistenceBinaryDataParts)
                {
                    binNew = NewPart.AddEmbeddedControlPersistenceBinaryDataPart(bin.ContentType, ExistingPart.GetIdOfPart(bin));
                    this.FeedDataEmbeddedControlPersistenceBinaryDataPart(binNew, bin);
                }
            }

            if (ExistingPart.EmbeddedControlPersistenceParts != null)
            {
                EmbeddedControlPersistencePart ecppNew;
                foreach (EmbeddedControlPersistencePart ecpp in ExistingPart.EmbeddedControlPersistenceParts)
                {
                    ecppNew = NewPart.AddEmbeddedControlPersistencePart(ecpp.ContentType, ExistingPart.GetIdOfPart(ecpp));
                    this.FeedDataEmbeddedControlPersistencePart(ecppNew, ecpp);
                }
            }

            if (ExistingPart.EmbeddedObjectParts != null)
            {
                EmbeddedObjectPart eopNew;
                foreach (EmbeddedObjectPart eop in ExistingPart.EmbeddedObjectParts)
                {
                    eopNew = NewPart.AddEmbeddedObjectPart(ExistingPart.GetIdOfPart(eop));
                    this.FeedDataEmbeddedObjectPart(eopNew, eop);
                }
            }

            if (ExistingPart.EmbeddedPackageParts != null)
            {
                EmbeddedPackagePart eppNew;
                foreach (EmbeddedPackagePart epp in ExistingPart.EmbeddedPackageParts)
                {
                    eppNew = NewPart.AddEmbeddedPackagePart(ExistingPart.GetIdOfPart(epp));
                    this.FeedDataEmbeddedPackagePart(eppNew, epp);
                }
            }

            foreach (ImagePart imgp in ExistingPart.ImageParts)
            {
                imgpNew = NewPart.AddImagePart(imgp.ContentType, ExistingPart.GetIdOfPart(imgp));
                this.FeedDataImagePart(imgpNew, imgp);
            }

            #region NotesSlidePart
            // TODO: there are a lot of stuff here...
            #endregion

            if (ExistingPart.SlideCommentsPart != null)
            {
                NewPart.AddNewPart<SlideCommentsPart>(ExistingPart.GetIdOfPart(ExistingPart.SlideCommentsPart));
                using (StreamReader sr = new StreamReader(ExistingPart.SlideCommentsPart.GetStream()))
                {
                    using (StreamWriter sw = new StreamWriter(NewPart.SlideCommentsPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(sr.ReadToEnd());
                    }
                }
            }

            #region SlideLayoutPart
            // TODO: there are a lot of stuff here...
            #endregion

            // what is going on? A SlidePart with SlideParts...
            // Hesitate to do this...
            if (ExistingPart.SlideParts != null)
            {
                SlidePart spNew;
                foreach (SlidePart sp in ExistingPart.SlideParts)
                {
                    spNew = NewPart.AddNewPart<SlidePart>(ExistingPart.GetIdOfPart(sp));
                    this.FeedDataSlidePart(spNew, sp);
                }
            }

            if (ExistingPart.SlideSyncDataPart != null)
            {
                NewPart.AddNewPart<SlideSyncDataPart>(ExistingPart.GetIdOfPart(ExistingPart.SlideSyncDataPart));
                using (StreamReader sr = new StreamReader(ExistingPart.SlideSyncDataPart.GetStream()))
                {
                    using (StreamWriter sw = new StreamWriter(NewPart.SlideSyncDataPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(sr.ReadToEnd());
                    }
                }
            }

            if (ExistingPart.ThemeOverridePart != null)
            {
                NewPart.AddNewPart<ThemeOverridePart>(ExistingPart.GetIdOfPart(ExistingPart.ThemeOverridePart));
                this.FeedDataThemeOverridePart(NewPart.ThemeOverridePart, ExistingPart.ThemeOverridePart);
            }

            if (ExistingPart.UserDefinedTagsParts != null)
            {
                UserDefinedTagsPart udtpNew;
                foreach (UserDefinedTagsPart udtp in ExistingPart.UserDefinedTagsParts)
                {
                    udtpNew = NewPart.AddNewPart<UserDefinedTagsPart>(ExistingPart.GetIdOfPart(udtp));
                    using (StreamReader sr = new StreamReader(udtp.GetStream()))
                    {
                        using (StreamWriter sw = new StreamWriter(udtpNew.GetStream(FileMode.Create)))
                        {
                            sw.Write(sr.ReadToEnd());
                        }
                    }
                }
            }

            if (ExistingPart.VmlDrawingParts != null)
            {
                VmlDrawingPart vdpNew;
                foreach (VmlDrawingPart vdp in ExistingPart.VmlDrawingParts)
                {
                    vdpNew = NewPart.AddNewPart<VmlDrawingPart>(ExistingPart.GetIdOfPart(vdp));
                    this.FeedDataVmlDrawingPart(vdpNew, vdp);
                }
            }
        }

        internal void FeedDataChartPart(ChartPart NewPart, ChartPart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            ImagePart imgpNew;

            if (ExistingPart.ChartDrawingPart != null)
            {
                NewPart.AddNewPart<ChartDrawingPart>(ExistingPart.GetIdOfPart(ExistingPart.ChartDrawingPart));
                using (StreamReader sr = new StreamReader(ExistingPart.ChartDrawingPart.GetStream()))
                {
                    using (StreamWriter sw = new StreamWriter(NewPart.ChartDrawingPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(sr.ReadToEnd());
                    }
                }

                // why does a ChartPart contain a ChartDrawingPart that contains a ChartPart??
                // Does it never end??
                if (ExistingPart.ChartDrawingPart.ChartPart != null)
                {
                    NewPart.ChartDrawingPart.AddNewPart<ChartPart>(ExistingPart.ChartDrawingPart.GetIdOfPart(ExistingPart.ChartDrawingPart.ChartPart));
                    this.FeedDataChartPart(NewPart.ChartDrawingPart.ChartPart, ExistingPart.ChartDrawingPart.ChartPart);
                }

                foreach (ImagePart imgp in ExistingPart.ChartDrawingPart.ImageParts)
                {
                    imgpNew = NewPart.ChartDrawingPart.AddImagePart(imgp.ContentType, ExistingPart.ChartDrawingPart.GetIdOfPart(imgp));
                    this.FeedDataImagePart(imgpNew, imgp);
                }
            }

            if (ExistingPart.EmbeddedPackagePart != null)
            {
                NewPart.AddEmbeddedPackagePart(ExistingPart.EmbeddedPackagePart.ContentType);
                this.FeedDataEmbeddedPackagePart(NewPart.EmbeddedPackagePart, ExistingPart.EmbeddedPackagePart);
            }

            foreach (ImagePart imgp in ExistingPart.ImageParts)
            {
                imgpNew = NewPart.AddImagePart(imgp.ContentType, ExistingPart.GetIdOfPart(imgp));
                this.FeedDataImagePart(imgpNew, imgp);
            }

            if (ExistingPart.ThemeOverridePart != null)
            {
                NewPart.AddNewPart<ThemeOverridePart>(ExistingPart.GetIdOfPart(ExistingPart.ThemeOverridePart));
                this.FeedDataThemeOverridePart(NewPart.ThemeOverridePart, ExistingPart.ThemeOverridePart);
            }
        }

        internal void FeedDataCustomXmlPart(CustomXmlPart NewPart, CustomXmlPart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            if (ExistingPart.CustomXmlPropertiesPart != null)
            {
                NewPart.AddNewPart<CustomXmlPropertiesPart>(ExistingPart.GetIdOfPart(ExistingPart.CustomXmlPropertiesPart));
                using (StreamReader sr = new StreamReader(ExistingPart.CustomXmlPropertiesPart.GetStream()))
                {
                    using (StreamWriter sw = new StreamWriter(NewPart.CustomXmlPropertiesPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(sr.ReadToEnd());
                    }
                }
            }
        }

        internal void FeedDataDiagramColorsPart(DiagramColorsPart NewPart, DiagramColorsPart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }
        }
        
        internal void FeedDataDiagramDataPart(DiagramDataPart NewPart, DiagramDataPart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            if (ExistingPart.ImageParts != null)
            {
                ImagePart imgpNew;
                foreach (ImagePart imgp in ExistingPart.ImageParts)
                {
                    imgpNew = NewPart.AddImagePart(imgp.ContentType, ExistingPart.GetIdOfPart(imgp));
                    this.FeedDataImagePart(imgpNew, imgp);
                }
            }

            if (ExistingPart.SlideParts != null)
            {
                SlidePart spNew;
                foreach (SlidePart sp in ExistingPart.SlideParts)
                {
                    spNew = NewPart.AddNewPart<SlidePart>(ExistingPart.GetIdOfPart(sp));
                    this.FeedDataSlidePart(spNew, sp);
                }
            }

            if (ExistingPart.WorksheetParts != null)
            {
                WorksheetPart wspNew;
                foreach (WorksheetPart wsp in ExistingPart.WorksheetParts)
                {
                    // will just do 1st level copying. NO NESTING!! That'll have to do...
                    wspNew = NewPart.AddNewPart<WorksheetPart>(ExistingPart.GetIdOfPart(wsp));
                    using (StreamReader sr = new StreamReader(wsp.GetStream()))
                    {
                        using (StreamWriter sw = new StreamWriter(wspNew.GetStream(FileMode.Create)))
                        {
                            sw.Write(sr.ReadToEnd());
                        }
                    }
                }
            }
        }

        internal void FeedDataDiagramLayoutDefinitionPart(DiagramLayoutDefinitionPart NewPart, DiagramLayoutDefinitionPart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            ImagePart imgpNew;
            foreach (ImagePart imgp in ExistingPart.ImageParts)
            {
                imgpNew = NewPart.AddImagePart(imgp.ContentType, ExistingPart.GetIdOfPart(imgp));
                this.FeedDataImagePart(imgpNew, imgp);
            }
        }

        internal void FeedDataDiagramPersistLayoutPart(DiagramPersistLayoutPart NewPart, DiagramPersistLayoutPart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            ImagePart imgpNew;
            foreach (ImagePart imgp in ExistingPart.ImageParts)
            {
                imgpNew = NewPart.AddImagePart(imgp.ContentType, ExistingPart.GetIdOfPart(imgp));
                this.FeedDataImagePart(imgpNew, imgp);
            }
        }

        internal void FeedDataDiagramStylePart(DiagramStylePart NewPart, DiagramStylePart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }
        }

        internal void FeedDataEmbeddedControlPersistencePart(EmbeddedControlPersistencePart NewPart, EmbeddedControlPersistencePart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            EmbeddedControlPersistenceBinaryDataPart binNew;
            foreach (EmbeddedControlPersistenceBinaryDataPart bin in ExistingPart.EmbeddedControlPersistenceBinaryDataParts)
            {
                binNew = NewPart.AddEmbeddedControlPersistenceBinaryDataPart(bin.ContentType, ExistingPart.GetIdOfPart(bin));
                this.FeedDataEmbeddedControlPersistenceBinaryDataPart(binNew, bin);
            }
        }

        internal void FeedDataEmbeddedControlPersistenceBinaryDataPart(EmbeddedControlPersistenceBinaryDataPart NewPart, EmbeddedControlPersistenceBinaryDataPart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }
        }

        internal void FeedDataEmbeddedObjectPart(EmbeddedObjectPart NewPart, EmbeddedObjectPart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }
        }

        internal void FeedDataEmbeddedPackagePart(EmbeddedPackagePart NewPart, EmbeddedPackagePart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }
        }

        internal void FeedDataThemeOverridePart(ThemeOverridePart NewPart, ThemeOverridePart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            ImagePart imgpNew;
            foreach (ImagePart imgp in ExistingPart.ImageParts)
            {
                imgpNew = NewPart.AddImagePart(imgp.ContentType, ExistingPart.GetIdOfPart(imgp));
                this.FeedDataImagePart(imgpNew, imgp);
            }
        }

        internal void FeedDataVmlDrawingPart(VmlDrawingPart NewPart, VmlDrawingPart ExistingPart)
        {
            using (StreamReader sr = new StreamReader(ExistingPart.GetStream()))
            {
                using (StreamWriter sw = new StreamWriter(NewPart.GetStream(FileMode.Create)))
                {
                    sw.Write(sr.ReadToEnd());
                }
            }

            ImagePart imgpNew;
            foreach (ImagePart imgp in ExistingPart.ImageParts)
            {
                imgpNew = NewPart.AddImagePart(imgp.ContentType, ExistingPart.GetIdOfPart(imgp));
                this.FeedDataImagePart(imgpNew, imgp);
            }

            LegacyDiagramTextPart ldtpNew;
            foreach (var ldtp in ExistingPart.LegacyDiagramTextParts)
            {
                ldtpNew = NewPart.AddNewPart<LegacyDiagramTextPart>(ExistingPart.GetIdOfPart(ldtp));
                using (StreamReader sr = new StreamReader(ldtp.GetStream()))
                {
                    using (StreamWriter sw = new StreamWriter(ldtpNew.GetStream(FileMode.Create)))
                    {
                        sw.Write(sr.ReadToEnd());
                    }
                }
            }
        }

        /// <summary>
        /// Move a worksheet to a new position in the spreadsheet.
        /// </summary>
        /// <param name="WorksheetName">The name of the worksheet.</param>
        /// <param name="Position">The new 1-based position index. Use 1 for 1st position, 2 for 2nd position and so on.</param>
        /// <returns>True if an actual move was done (this excludes when the given worksheet is already in the given position). False otherwise.</returns>
        public bool MoveWorksheet(string WorksheetName, int Position)
        {
            bool result = false;
            // get the zero-based index first
            Position = Position - 1;

            if (Position < 0) Position = 0;
            if (Position >= slwb.Sheets.Count) Position = slwb.Sheets.Count - 1;

            // orig order, SheetId
            Dictionary<uint, uint> OrigOrder = new Dictionary<uint, uint>();
            // SheetId, new order
            Dictionary<uint, uint> NewOrder = new Dictionary<uint, uint>();

            bool bFound = false;
            SLSheet sheet = slwb.Sheets[0];
            int iCurrentPosition = 0;
            int i = 0;
            for (i = 0; i < slwb.Sheets.Count; ++i)
            {
                OrigOrder[(uint)i] = slwb.Sheets[i].SheetId;
                if (slwb.Sheets[i].Name.Equals(WorksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    iCurrentPosition = i;
                    sheet = slwb.Sheets[i];
                    bFound = true;
                }
            }

            if (iCurrentPosition != Position && bFound)
            {
                slwb.Sheets.RemoveAt(iCurrentPosition);
                slwb.Sheets.Insert(Position, sheet);

                for (i = 0; i < slwb.Sheets.Count; ++i)
                {
                    NewOrder[slwb.Sheets[i].SheetId] = (uint)i;
                }

                foreach (SLDefinedName dn in slwb.DefinedNames)
                {
                    if (dn.LocalSheetId != null)
                    {
                        if (OrigOrder.ContainsKey(dn.LocalSheetId.Value))
                        {
                            if (NewOrder.ContainsKey(OrigOrder[dn.LocalSheetId.Value]))
                            {
                                dn.LocalSheetId = NewOrder[OrigOrder[dn.LocalSheetId.Value]];
                            }
                        }
                    }
                }

                result = true;
            }

            return result;
        }

        /// <summary>
        /// Set the default row height for the currently selected worksheet.
        /// </summary>
        /// <param name="RowHeight">The row height in points.</param>
        public void SetWorksheetDefaultRowHeight(double RowHeight)
        {
            slws.SheetFormatProperties.DefaultRowHeight = RowHeight;
            slws.SheetFormatProperties.CustomHeight = true;

            // TODO: resize images and charts
        }

        /// <summary>
        /// Set the default column width for the currently selected worksheet.
        /// </summary>
        /// <param name="ColumnWidth">The column width.</param>
        public void SetWorksheetDefaultColumnWidth(double ColumnWidth)
        {
            slws.SheetFormatProperties.DefaultColumnWidth = ColumnWidth;
            // TODO: resize images and charts
        }

        /// <summary>
        /// Freeze panes in the worksheet (for the first workbook view). Will do nothing if both parameters are zero (because there's nothing to freeze). Will also do nothing if either of the parameters is equal to their respective limits (maximum number of rows, or maximum number of columns).
        /// </summary>
        /// <param name="NumberOfTopMostRows">Number of top-most rows to keep in place.</param>
        /// <param name="NumberOfLeftMostColumns">Number of left-most columns to keep in place.</param>
        public void FreezePanes(int NumberOfTopMostRows, int NumberOfLeftMostColumns)
        {
            if (NumberOfTopMostRows == 0 && NumberOfLeftMostColumns == 0) return;
            // no point if they're negative, right?
            if (NumberOfTopMostRows < 0 || NumberOfLeftMostColumns < 0) return;
            // the "greater than" part ensures correct limit checks as well.
            if (NumberOfTopMostRows >= SLConstants.RowLimit || NumberOfLeftMostColumns >= SLConstants.ColumnLimit) return;

            SLSheetView slsv = new SLSheetView();
            // will use the first workbook view by default
            slsv.WorkbookViewId = 0;
            if (NumberOfLeftMostColumns > 0) slsv.Pane.HorizontalSplit = NumberOfLeftMostColumns;
            if (NumberOfTopMostRows > 0) slsv.Pane.VerticalSplit = NumberOfTopMostRows;

            int iRowIndex = NumberOfTopMostRows + 1;
            int iColumnIndex = NumberOfLeftMostColumns + 1;

            SLSelection sel;

            slsv.Pane.TopLeftCell = SLTool.ToCellReference(iRowIndex, iColumnIndex);

            if (NumberOfLeftMostColumns == 0)
            {
                slsv.Pane.ActivePane = PaneValues.BottomLeft;
                slsv.Pane.State = PaneStateValues.Frozen;
                slsv.Pane.VerticalSplit = NumberOfTopMostRows;

                sel = new SLSelection();
                sel.Pane = PaneValues.BottomLeft;
                sel.ActiveCell = SLTool.ToCellReference(iRowIndex, iColumnIndex);
                sel.SequenceOfReferences.Add(new SLCellPointRange(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex));
                slsv.Selections.Add(sel);
            }
            else if (NumberOfTopMostRows == 0)
            {
                slsv.Pane.ActivePane = PaneValues.TopRight;
                slsv.Pane.State = PaneStateValues.Frozen;
                slsv.Pane.HorizontalSplit = NumberOfLeftMostColumns;

                sel = new SLSelection();
                sel.Pane = PaneValues.TopRight;
                sel.ActiveCell = SLTool.ToCellReference(iRowIndex, iColumnIndex);
                sel.SequenceOfReferences.Add(new SLCellPointRange(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex));
                slsv.Selections.Add(sel);
            }
            else
            {
                slsv.Pane.ActivePane = PaneValues.BottomRight;
                slsv.Pane.State = PaneStateValues.Frozen;

                sel = new SLSelection();
                sel.Pane = PaneValues.TopRight;
                sel.ActiveCell = SLTool.ToCellReference(1, iColumnIndex);
                sel.SequenceOfReferences.Add(new SLCellPointRange(1, iColumnIndex, 1, iColumnIndex));
                slsv.Selections.Add(sel);

                sel = new SLSelection();
                sel.Pane = PaneValues.BottomLeft;
                sel.ActiveCell = SLTool.ToCellReference(iRowIndex, 1);
                sel.SequenceOfReferences.Add(new SLCellPointRange(iRowIndex, 1, iRowIndex, 1));
                slsv.Selections.Add(sel);

                sel = new SLSelection();
                sel.Pane = PaneValues.BottomRight;
                if (slws.ActiveCell.RowIndex == 1 && slws.ActiveCell.ColumnIndex == 1)
                {
                    sel.ActiveCell = SLTool.ToCellReference(iRowIndex, iColumnIndex);
                    sel.SequenceOfReferences.Add(new SLCellPointRange(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex));
                }
                else
                {
                    sel.ActiveCell = SLTool.ToCellReference(slws.ActiveCell.RowIndex, slws.ActiveCell.ColumnIndex);
                    sel.SequenceOfReferences.Add(new SLCellPointRange(slws.ActiveCell.RowIndex, slws.ActiveCell.ColumnIndex, slws.ActiveCell.RowIndex, slws.ActiveCell.ColumnIndex));
                }
                slsv.Selections.Add(sel);
            }

            bool bFound = false;
            foreach (SLSheetView sv in slws.SheetViews)
            {
                if (sv.WorkbookViewId == 0)
                {
                    bFound = true;
                    sv.Pane = slsv.Pane.Clone();

                    sv.Selections = new List<SLSelection>();
                    foreach (SLSelection slsel in slsv.Selections)
                    {
                        sv.Selections.Add(slsel.Clone());
                    }

                    sv.PivotSelections = new List<PivotSelection>();
                }
            }

            if (!bFound)
            {
                slws.SheetViews.Add(slsv);
            }
        }

        /// <summary>
        /// Unfreeze the frozen panes in the worksheet (for the first workbook view).
        /// </summary>
        public void UnfreezePanes()
        {
            this.UnfreezeUnsplitPanes(true);
        }

        /// <summary>
        /// Split panes in the worksheet (for the first workbook view). Will do nothing if both number of rows and number of columns are zero (because there's nothing to split). Will also do nothing if either is equal to their respective limits (maximum number of rows, or maximum number of columns).
        /// </summary>
        /// <param name="NumberOfRows">Number of top-most rows above the horizontal split line.</param>
        /// <param name="NumberOfColumns">Number of left-most columns left of the vertical split line.</param>
        /// <param name="ShowRowColumnHeadings">True if the row and column headings are shown. False otherwise.</param>
        public void SplitPanes(int NumberOfRows, int NumberOfColumns, bool ShowRowColumnHeadings)
        {
            if (!ShowRowColumnHeadings)
            {
                this.SplitPanes(NumberOfRows, NumberOfColumns, false, 0, 0, false);
            }
            else
            {
                this.SplitPanes(NumberOfRows, NumberOfColumns, true, slws.SheetFormatProperties.DefaultRowHeight, SLTool.GetDefaultRowHeadingWidth(SimpleTheme.MinorLatinFont, gbThrowExceptionsIfAny), false);
            }
        }

        /// <summary>
        /// Split panes in the worksheet (for the first workbook view). Will do nothing if both number of rows and number of columns are zero (because there's nothing to split). Will also do nothing if either is equal to their respective limits (maximum number of rows, or maximum number of columns).
        /// </summary>
        /// <param name="NumberOfRows">Number of top-most rows above the horizontal split line.</param>
        /// <param name="NumberOfColumns">Number of left-most columns left of the vertical split line.</param>
        /// <param name="ShowRowColumnHeadings">True if the row and column headings are shown. False otherwise.</param>
        /// <param name="VerticalOffsetInPoints">This is more useful when row and column headings are shown. This will be the height of the column heading in points.</param>
        /// <param name="HorizontalOffsetInPoints">This is more useful when row and column headings are shown. This will be the width of the row heading in points.</param>
        public void SplitPanes(int NumberOfRows, int NumberOfColumns, bool ShowRowColumnHeadings, double VerticalOffsetInPoints, double HorizontalOffsetInPoints)
        {
            this.SplitPanes(NumberOfRows, NumberOfColumns, ShowRowColumnHeadings, VerticalOffsetInPoints, HorizontalOffsetInPoints, true);
        }

        /// <summary>
        /// Split panes in the worksheet (for the first workbook view). Will do nothing if both number of rows and number of columns are zero (because there's nothing to split). Will also do nothing if either is equal to their respective limits (maximum number of rows, or maximum number of columns).
        /// The underlying engine tries to guess the individual row heights and column widths. Then the horizontal and vertical split lines are placed based on the guesses.
        /// Forcing the row and column dimensions to fit the split lines might mean the worksheet looking oddly sized.
        /// </summary>
        /// <param name="NumberOfRows">Number of top-most rows above the horizontal split line.</param>
        /// <param name="NumberOfColumns">Number of left-most columns left of the vertical split line.</param>
        /// <param name="ShowRowColumnHeadings">True if the row and column headings are shown. False otherwise.</param>
        /// <param name="VerticalOffsetInPoints">This is more useful when row and column headings are shown. This will be the height of the column heading in points.</param>
        /// <param name="HorizontalOffsetInPoints">This is more useful when row and column headings are shown. This will be the width of the row heading in points.</param>
        /// <param name="ForceCustomRowColumnDimensions">Set true to force the worksheet's row height and column width to fit the given horizontal and vertical splits. False otherwise.</param>
        public void SplitPanes(int NumberOfRows, int NumberOfColumns, bool ShowRowColumnHeadings, double VerticalOffsetInPoints, double HorizontalOffsetInPoints, bool ForceCustomRowColumnDimensions)
        {
            // At Calibri 11pt as minor font at 96 DPI,
            // suggested vertical offset is 15 points ~= 20 pixels high
            // suggested horizontal offset is 19.5 points ~= 26 pixels wide
            // At Calibri 11pt as minor font at 120 DPI,
            // suggested vertical offset is 14.4 points ~= 24 pixels high
            // suggested horizontal offset is 20.4 points ~= 34 pixels wide

            if (NumberOfRows == 0 && NumberOfColumns == 0) return;
            // no point if they're negative, right?
            if (NumberOfRows < 0 || NumberOfColumns < 0) return;
            // the "greater than" part ensures correct limit checks as well.
            if (NumberOfRows >= SLConstants.RowLimit || NumberOfColumns >= SLConstants.ColumnLimit) return;

            // we don't care for negative offsets...
            if (VerticalOffsetInPoints < 0) VerticalOffsetInPoints = 0;
            if (HorizontalOffsetInPoints < 0) HorizontalOffsetInPoints = 0;

            slws.ForceCustomRowColumnDimensionsSplitting = ForceCustomRowColumnDimensions;

            long lWidth = 0, lHeight = 0;
            int i = 0;

            SLSheetView slsv = new SLSheetView();
            // will use the first workbook view by default
            slsv.WorkbookViewId = 0;
            slsv.ShowRowColHeaders = ShowRowColumnHeadings;

            if (NumberOfColumns > 0)
            {
                SLColumnProperties cp;
                for (i = 1; i <= NumberOfColumns; ++i)
                {
                    if (slws.ColumnProperties.ContainsKey(i))
                    {
                        cp = slws.ColumnProperties[i];
                        if (cp.HasWidth)
                        {
                            lWidth += cp.WidthInEMU;
                        }
                        else
                        {
                            lWidth += slws.SheetFormatProperties.DefaultColumnWidthInEMU;
                        }
                    }
                    else
                    {
                        lWidth += slws.SheetFormatProperties.DefaultColumnWidthInEMU;
                    }
                }
            }
            // split lengths are in twentieth's of a point
            if (ShowRowColumnHeadings)
            {
                lWidth = (long)Math.Round(20.0 * (((double)lWidth / (double)SLConstants.PointToEMU) + HorizontalOffsetInPoints));
            }
            else
            {
                lWidth = (long)Math.Round(20.0 * ((double)lWidth / (double)SLConstants.PointToEMU));
            }
            slsv.Pane.HorizontalSplit = lWidth;

            if (NumberOfRows > 0)
            {
                SLRowProperties rp;
                for (i = 1; i <= NumberOfRows; ++i)
                {
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        if (rp.HasHeight)
                        {
                            lHeight += rp.HeightInEMU;
                        }
                        else
                        {
                            lHeight += slws.SheetFormatProperties.DefaultRowHeightInEMU;
                        }
                    }
                    else
                    {
                        lHeight += slws.SheetFormatProperties.DefaultRowHeightInEMU;
                    }
                }
            }
            // split lengths are in twentieth's of a point
            if (ShowRowColumnHeadings)
            {
                lHeight = (long)Math.Round(20.0 * (((double)lHeight / (double)SLConstants.PointToEMU) + VerticalOffsetInPoints));
            }
            else
            {
                lHeight = (long)Math.Round(20.0 * ((double)lHeight / (double)SLConstants.PointToEMU));
            }
            slsv.Pane.VerticalSplit = lHeight;

            int iRowIndex = NumberOfRows + 1;
            int iColumnIndex = NumberOfColumns + 1;

            slsv.Pane.TopLeftCell = SLTool.ToCellReference(iRowIndex, iColumnIndex);
            slsv.Pane.ActivePane = PaneValues.BottomRight;
            slsv.Pane.State = PaneStateValues.Split;

            SLSelection sel;

            if (slws.ActiveCell.RowIndex < iRowIndex)
            {
                if (slws.ActiveCell.ColumnIndex < iColumnIndex)
                {
                    // it seems that if the active cell is A1 Excel don't render the selection XML tag.
                    if (slws.ActiveCell.RowIndex != 1 || slws.ActiveCell.ColumnIndex != 1)
                    {
                        slsv.Pane.ActivePane = PaneValues.TopLeft;

                        sel = new SLSelection();
                        sel.ActiveCell = SLTool.ToCellReference(slws.ActiveCell.RowIndex, slws.ActiveCell.ColumnIndex);
                        sel.SequenceOfReferences.Add(new SLCellPointRange(slws.ActiveCell.RowIndex, slws.ActiveCell.ColumnIndex, slws.ActiveCell.RowIndex, slws.ActiveCell.ColumnIndex));
                        slsv.Selections.Add(sel);
                    }
                }
                else
                {
                    slsv.Pane.ActivePane = PaneValues.TopRight;
                }
            }
            else
            {
                if (slws.ActiveCell.ColumnIndex < iColumnIndex)
                {
                    slsv.Pane.ActivePane = PaneValues.BottomLeft;
                }
                else
                {
                    slsv.Pane.ActivePane = PaneValues.BottomRight;
                }
            }

            sel = new SLSelection();
            sel.Pane = PaneValues.TopRight;
            sel.ActiveCell = SLTool.ToCellReference(1, iColumnIndex);
            sel.SequenceOfReferences.Add(new SLCellPointRange(1, iColumnIndex, 1, iColumnIndex));
            slsv.Selections.Add(sel);

            sel = new SLSelection();
            sel.Pane = PaneValues.BottomLeft;
            sel.ActiveCell = SLTool.ToCellReference(iRowIndex, 1);
            sel.SequenceOfReferences.Add(new SLCellPointRange(iRowIndex, 1, iRowIndex, 1));
            slsv.Selections.Add(sel);

            sel = new SLSelection();
            sel.Pane = PaneValues.BottomRight;
            sel.ActiveCell = SLTool.ToCellReference(iRowIndex, iColumnIndex);
            sel.SequenceOfReferences.Add(new SLCellPointRange(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex));
            slsv.Selections.Add(sel);

            bool bFound = false;
            foreach (SLSheetView sv in slws.SheetViews)
            {
                if (sv.WorkbookViewId == 0)
                {
                    bFound = true;
                    sv.Pane = slsv.Pane.Clone();

                    sv.Selections = new List<SLSelection>();
                    foreach (SLSelection slsel in slsv.Selections)
                    {
                        sv.Selections.Add(slsel.Clone());
                    }

                    sv.PivotSelections = new List<PivotSelection>();
                }
            }

            if (!bFound)
            {
                slws.SheetViews.Add(slsv);
            }
        }

        /// <summary>
        /// Unsplit the split panes in the worksheet (for the first workbook view).
        /// </summary>
        public void UnsplitPanes()
        {
            this.UnfreezeUnsplitPanes(false);
        }

        private void UnfreezeUnsplitPanes(bool IsFreeze)
        {
            bool bToRemove = false;
            foreach (SLSheetView sv in slws.SheetViews)
            {
                if (sv.WorkbookViewId == 0)
                {
                    if (IsFreeze && sv.Pane.State == PaneStateValues.Frozen)
                    {
                        bToRemove = true;
                    }
                    else if (!IsFreeze && sv.Pane.State == PaneStateValues.Split)
                    {
                        bToRemove = true;
                    }

                    if (bToRemove)
                    {
                        sv.Pane = new SLPane();
                        sv.Selections = new List<SLSelection>();
                        sv.PivotSelections = new List<PivotSelection>();
                    }

                    break;
                }
            }
        }

        /// <summary>
        /// Add a background picture to the currently selected worksheet given the file name of a picture.
        /// If there's an existing background picture, that will be deleted first.
        /// </summary>
        /// <param name="FileName">The file name of a picture to be used.</param>
        public void AddBackgroundPicture(string FileName)
        {
            // delete any background picture first
            this.DeleteBackgroundPicture();

            slws.BackgroundPictureDataIsInFile = true;
            slws.BackgroundPictureFileName = FileName;
            slws.BackgroundPictureImagePartType = SLA.SLDrawingTool.GetImagePartType(FileName);
        }

        /// <summary>
        /// Add a background picture to the currently selected worksheet given a picture's data in a byte array.
        /// If there's an existing background picture, that will be deleted first.
        /// </summary>
        /// <param name="PictureByteData">The picture's data in a byte array.</param>
        /// <param name="PictureType">The image type of the picture.</param>
        public void AddBackgroundPicture(byte[] PictureByteData, ImagePartType PictureType)
        {
            // delete any background picture first
            this.DeleteBackgroundPicture();

            slws.BackgroundPictureDataIsInFile = false;
            slws.BackgroundPictureByteData = new byte[PictureByteData.Length];
            for (int i = 0; i < PictureByteData.Length; ++i)
            {
                slws.BackgroundPictureByteData[i] = PictureByteData[i];
            }
            slws.BackgroundPictureImagePartType = PictureType;
        }

        /// <summary>
        /// Delete the background picture of the currently selected worksheet.
        /// </summary>
        public void DeleteBackgroundPicture()
        {
            if (slws.BackgroundPictureId.Length > 0)
            {
                if (!IsNewWorksheet)
                {
                    if (!string.IsNullOrEmpty(gsSelectedWorksheetRelationshipID))
                    {
                        WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(gsSelectedWorksheetRelationshipID);
                        wsp.DeletePart(slws.BackgroundPictureId);
                    }
                }
            }

            slws.InitializeBackgroundPictureStuff();
        }

        /// <summary>
        /// Insert a picture into the currently selected worksheet.
        /// </summary>
        /// <param name="Picture">An SLPicture object with the picture's properties already set.</param>
        /// <returns>True if the picture is successfully inserted. False otherwise.</returns>
        public bool InsertPicture(Drawing.SLPicture Picture)
        {
            if (Picture.UseEasyPositioning)
            {
                int AnchorRowIndex = 0;
                long AnchorRowOffset = 0;
                int AnchorColumnIndex = 0;
                long AnchorColumnOffset = 0;
                double fTemp = 0;
                SLRowProperties rp;
                SLColumnProperties cp;

                AnchorRowIndex = (int)Math.Floor(Picture.TopPosition);
                fTemp = Picture.TopPosition - AnchorRowIndex;
                AnchorRowOffset = (long)(fTemp * slws.SheetFormatProperties.DefaultRowHeightInEMU);
                ++AnchorRowIndex;
                if (slws.RowProperties.ContainsKey(AnchorRowIndex))
                {
                    rp = slws.RowProperties[AnchorRowIndex];
                    if (rp.HasHeight) AnchorRowOffset = (long)(fTemp * rp.HeightInEMU);
                }

                AnchorColumnIndex = (int)Math.Floor(Picture.LeftPosition);
                fTemp = Picture.LeftPosition - AnchorColumnIndex;

                AnchorColumnOffset = (long)(fTemp * slws.SheetFormatProperties.DefaultColumnWidthInEMU);

                ++AnchorColumnIndex;
                if (slws.ColumnProperties.ContainsKey(AnchorColumnIndex))
                {
                    cp = slws.ColumnProperties[AnchorColumnIndex];
                    if (cp.HasWidth)
                    {
                        AnchorColumnOffset = (long)(fTemp * cp.WidthInEMU);
                    }
                }

                Picture.AnchorRowIndex = AnchorRowIndex;
                Picture.AnchorColumnIndex = AnchorColumnIndex;
                Picture.OffsetX = AnchorColumnOffset;
                Picture.OffsetY = AnchorRowOffset;
            }

            if (!SLTool.CheckRowColumnIndexLimit(Picture.AnchorRowIndex, Picture.AnchorColumnIndex))
            {
                return false;
            }

            slws.Pictures.Add(Picture.Clone());

            return true;
        }

        /// <summary>
        /// Insert a sparkline group into the currently selected worksheet. If unsuccessful, please check that your sparkline location is correctly set. See SetLocation() for details.
        /// </summary>
        /// <param name="SparklineGroup">An SLSparklineGroup object with the properties already set.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertSparklineGroup(SLSparklineGroup SparklineGroup)
        {
            if (SparklineGroup.Sparklines.Count == 0)
            {
                return false;
            }
            else
            {
                slws.SparklineGroups.Add(SparklineGroup.Clone());
                return true;
            }
        }

        /// <summary>
        /// Clear all sparkline groups in the currently selected worksheet.
        /// </summary>
        public void ClearAllSparklineGroups()
        {
            slws.SparklineGroups.Clear();
        }

        /// <summary>
        /// Insert a chart into the currently selected worksheet.
        /// </summary>
        /// <param name="Chart">An SLChart object with the chart's properties already set.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertChart(Charts.SLChart Chart)
        {
            slws.Charts.Add(Chart.Clone());
            return true;
        }

        /// <summary>
        /// Insert a chart into a chartsheet.
        /// </summary>
        /// <param name="Chart">An SLChart object with the chart's properties already set.</param>
        /// <param name="ChartsheetName">The name should not be blank, nor exceed 31 characters. And it cannot contain these characters: \/?*[] It cannot be the same as an existing name (case-insensitive). But there's nothing stopping you from using 3 spaces as a name.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool InsertChart(Charts.SLChart Chart, string ChartsheetName)
        {
            if (!SLTool.CheckSheetChartName(ChartsheetName))
            {
                return false;
            }
            foreach (SLSheet sheet in slwb.Sheets)
            {
                if (sheet.Name.Equals(ChartsheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            ChartsheetPart csp = wbp.AddNewPart<ChartsheetPart>();

            #region the drawing part
            DrawingsPart dp = csp.AddNewPart<DrawingsPart>();
            dp.WorksheetDrawing = new Xdr.WorksheetDrawing();

            ChartPart chartp = dp.AddNewPart<ChartPart>();
            chartp.ChartSpace = Chart.ToChartSpace(ref chartp);

            Xdr.AbsoluteAnchor absanchor = new Xdr.AbsoluteAnchor();
            absanchor.Append(new Xdr.Position() { X = 0, Y = 0 });
            // The paper size is involved. The default is Letter, which is 11 inch by 8.5 inch
            // with landscape orientation.
            // The page margin is also involved. In particular, the left and right margins for
            // the width, and top and bottom margins for the height.
            // We're using the "default" margin settings, which is
            // left=0.7, right=0.7, top=0.75, bottom=0.75
            // Excel also has a buffer of 0.12 inches around the sides...
            // And so,
            // 8668512 = (11 - 0.7 - 0.7 - 0.12) * 914400
            // 6291072 = (8.5 - 0.75 - 0.75 - 0.12) * 914400
            absanchor.Append(new Xdr.Extent() { Cx = 8668512, Cy = 6291072 });

            Xdr.GraphicFrame gf = new Xdr.GraphicFrame() { Macro = "" };

            gf.NonVisualGraphicFrameProperties = new Xdr.NonVisualGraphicFrameProperties();
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties()
            {
                Id = 2,
                Name = Chart.ChartName
            };
            gf.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties = new Xdr.NonVisualGraphicFrameDrawingProperties()
            {
                GraphicFrameLocks = new A.GraphicFrameLocks() { NoGrouping = true }
            };

            gf.Transform = new Xdr.Transform()
            {
                Offset = new A.Offset() { X = 0, Y = 0 },
                Extents = new A.Extents() { Cx = 0, Cy = 0 }
            };

            gf.Graphic = new A.Graphic();
            gf.Graphic.GraphicData = new A.GraphicData();
            gf.Graphic.GraphicData.Uri = SLConstants.NamespaceC;
            gf.Graphic.GraphicData.Append(new C.ChartReference() { Id = dp.GetIdOfPart(chartp) });

            absanchor.Append(gf);
            absanchor.Append(new Xdr.ClientData());

            dp.WorksheetDrawing.Append(absanchor);
            #endregion

            #region the chartsheet part
            csp.Chartsheet = new Chartsheet();
            csp.Chartsheet.ChartSheetProperties = new ChartSheetProperties();
            csp.Chartsheet.ChartSheetViews = new ChartSheetViews();
            csp.Chartsheet.ChartSheetViews.Append(new ChartSheetView() { WorkbookViewId = 0 });
            csp.Chartsheet.PageMargins = new PageMargins()
            {
                Left = SLConstants.NormalLeftMargin,
                Right = SLConstants.NormalRightMargin,
                Top = SLConstants.NormalTopMargin,
                Bottom = SLConstants.NormalBottomMargin,
                Header = SLConstants.NormalHeaderMargin,
                Footer = SLConstants.NormalFooterMargin
            };
            csp.Chartsheet.Drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing()
            {
                Id = csp.GetIdOfPart(dp)
            };
            #endregion

            ++giWorksheetIdCounter;
            slwb.Sheets.Add(new SLSheet(ChartsheetName, (uint)giWorksheetIdCounter, wbp.GetIdOfPart(csp), SLSheetType.Chartsheet));

            return true;
        }

        private void WriteImageParts(DrawingsPart dp)
        {
            ImagePart imgp;
            Xdr.WorksheetDrawing wsd = dp.WorksheetDrawing;
            SLRowProperties rp;
            SLColumnProperties cp;

            #region Charts
            if (slws.Charts.Count > 0)
            {
                int FromAnchorRowIndex = 0;
                long FromAnchorRowOffset = 0;
                int FromAnchorColumnIndex = 0;
                long FromAnchorColumnOffset = 0;
                int ToAnchorRowIndex = 4;
                long ToAnchorRowOffset = 0;
                int ToAnchorColumnIndex = 4;
                long ToAnchorColumnOffset = 0;
                double fTemp = 0;

                ChartPart chartp;

                Xdr.GraphicFrame gf;

                foreach (Charts.SLChart Chart in slws.Charts)
                {
                    chartp = dp.AddNewPart<ChartPart>();
                    chartp.ChartSpace = Chart.ToChartSpace(ref chartp);

                    gf = new Xdr.GraphicFrame();
                    gf.Macro = string.Empty;
                    gf.NonVisualGraphicFrameProperties = new Xdr.NonVisualGraphicFrameProperties();
                    gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties();
                    gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Id = slws.NextWorksheetDrawingId;
                    ++slws.NextWorksheetDrawingId;
                    gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name = Chart.ChartName;
                    // alt text for charts
                    //...gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Description = "";
                    gf.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties = new Xdr.NonVisualGraphicFrameDrawingProperties();

                    gf.Transform = new Xdr.Transform();
                    gf.Transform.Offset = new A.Offset() { X = 0, Y = 0 };
                    gf.Transform.Extents = new A.Extents() { Cx = 0, Cy = 0 };

                    gf.Graphic = new A.Graphic();
                    gf.Graphic.GraphicData = new A.GraphicData();
                    gf.Graphic.GraphicData.Uri = SLConstants.NamespaceC;
                    gf.Graphic.GraphicData.Append(new C.ChartReference() { Id = dp.GetIdOfPart(chartp) });

                    FromAnchorRowIndex = (int)Math.Floor(Chart.TopPosition);
                    fTemp = Chart.TopPosition - FromAnchorRowIndex;
                    FromAnchorRowOffset = (long)(fTemp * slws.SheetFormatProperties.DefaultRowHeightInEMU);
                    if (slws.RowProperties.ContainsKey(FromAnchorRowIndex + 1))
                    {
                        rp = slws.RowProperties[FromAnchorRowIndex + 1];
                        if (rp.HasHeight) FromAnchorRowOffset = (long)(fTemp * rp.HeightInEMU);
                    }

                    FromAnchorColumnIndex = (int)Math.Floor(Chart.LeftPosition);
                    fTemp = Chart.LeftPosition - FromAnchorColumnIndex;

                    FromAnchorColumnOffset = (long)(fTemp * slws.SheetFormatProperties.DefaultColumnWidthInEMU);

                    if (slws.ColumnProperties.ContainsKey(FromAnchorColumnIndex + 1))
                    {
                        cp = slws.ColumnProperties[FromAnchorColumnIndex + 1];
                        if (cp.HasWidth)
                        {
                            FromAnchorColumnOffset = (long)(fTemp * cp.WidthInEMU);
                        }
                    }

                    ToAnchorRowIndex = (int)Math.Floor(Chart.BottomPosition);
                    fTemp = Chart.BottomPosition - ToAnchorRowIndex;
                    ToAnchorRowOffset = (long)(fTemp * slws.SheetFormatProperties.DefaultRowHeightInEMU);
                    if (slws.RowProperties.ContainsKey(ToAnchorRowIndex + 1))
                    {
                        rp = slws.RowProperties[ToAnchorRowIndex + 1];
                        if (rp.HasHeight) ToAnchorRowOffset = (long)(fTemp * rp.HeightInEMU);
                    }

                    ToAnchorColumnIndex = (int)Math.Floor(Chart.RightPosition);
                    fTemp = Chart.RightPosition - ToAnchorColumnIndex;

                    ToAnchorColumnOffset = (long)(fTemp * slws.SheetFormatProperties.DefaultColumnWidthInEMU);

                    if (slws.ColumnProperties.ContainsKey(ToAnchorColumnIndex + 1))
                    {
                        cp = slws.ColumnProperties[ToAnchorColumnIndex + 1];
                        if (cp.HasWidth)
                        {
                            ToAnchorColumnOffset = (long)(fTemp * cp.WidthInEMU);
                        }
                    }

                    Xdr.TwoCellAnchor tcanchor = new Xdr.TwoCellAnchor();
                    tcanchor.FromMarker = new Xdr.FromMarker();
                    tcanchor.FromMarker.RowId = new Xdr.RowId(FromAnchorRowIndex.ToString(CultureInfo.InvariantCulture));
                    tcanchor.FromMarker.RowOffset = new Xdr.RowOffset(FromAnchorRowOffset.ToString(CultureInfo.InvariantCulture));
                    tcanchor.FromMarker.ColumnId = new Xdr.ColumnId(FromAnchorColumnIndex.ToString(CultureInfo.InvariantCulture));
                    tcanchor.FromMarker.ColumnOffset = new Xdr.ColumnOffset(FromAnchorColumnOffset.ToString(CultureInfo.InvariantCulture));

                    tcanchor.ToMarker = new Xdr.ToMarker();
                    tcanchor.ToMarker.RowId = new Xdr.RowId(ToAnchorRowIndex.ToString(CultureInfo.InvariantCulture));
                    tcanchor.ToMarker.RowOffset = new Xdr.RowOffset(ToAnchorRowOffset.ToString(CultureInfo.InvariantCulture));
                    tcanchor.ToMarker.ColumnId = new Xdr.ColumnId(ToAnchorColumnIndex.ToString(CultureInfo.InvariantCulture));
                    tcanchor.ToMarker.ColumnOffset = new Xdr.ColumnOffset(ToAnchorColumnOffset.ToString(CultureInfo.InvariantCulture));

                    tcanchor.Append(gf);
                    tcanchor.Append(new Xdr.ClientData());

                    wsd.Append(tcanchor);
                    wsd.Save(dp);
                }
            }
            #endregion

            #region Pictures
            if (slws.Pictures.Count > 0)
            {
                foreach (Drawing.SLPicture Picture in slws.Pictures)
                {
                    imgp = dp.AddImagePart(Picture.PictureImagePartType);

                    if (Picture.DataIsInFile)
                    {
                        using (FileStream fs = new FileStream(Picture.PictureFileName, FileMode.Open))
                        {
                            imgp.FeedData(fs);
                        }
                    }
                    else
                    {
                        using (MemoryStream ms = new MemoryStream(Picture.PictureByteData))
                        {
                            imgp.FeedData(ms);
                        }
                    }

                    Xdr.Picture pic = new Xdr.Picture();
                    pic.NonVisualPictureProperties = new Xdr.NonVisualPictureProperties();

                    pic.NonVisualPictureProperties.NonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties();
                    pic.NonVisualPictureProperties.NonVisualDrawingProperties.Id = slws.NextWorksheetDrawingId;
                    ++slws.NextWorksheetDrawingId;
                    // recommendation is to set as the actual filename, but we'll follow Excel here...
                    // Note: the name value can be used multiple times without Excel choking.
                    // So for example, you can have two pictures with "Picture 1".
                    pic.NonVisualPictureProperties.NonVisualDrawingProperties.Name = string.Format("Picture {0}", dp.ImageParts.Count());
                    pic.NonVisualPictureProperties.NonVisualDrawingProperties.Description = Picture.AlternativeText;
                    // hlinkClick and hlinkHover as children

                    if (Picture.HasUri)
                    {
                        HyperlinkRelationship hlinkrel = dp.AddHyperlinkRelationship(new System.Uri(Picture.HyperlinkUri, Picture.HyperlinkUriKind), Picture.IsHyperlinkExternal);
                        pic.NonVisualPictureProperties.NonVisualDrawingProperties.HyperlinkOnClick = new A.HyperlinkOnClick() { Id = hlinkrel.Id };
                    }

                    pic.NonVisualPictureProperties.NonVisualPictureDrawingProperties = new Xdr.NonVisualPictureDrawingProperties();

                    pic.BlipFill = new Xdr.BlipFill();
                    pic.BlipFill.Blip = new A.Blip();
                    pic.BlipFill.Blip.Embed = dp.GetIdOfPart(imgp);
                    if (Picture.CompressionState != A.BlipCompressionValues.None)
                    {
                        pic.BlipFill.Blip.CompressionState = Picture.CompressionState;
                    }

                    if (Picture.Brightness != 0 || Picture.Contrast != 0)
                    {
                        A.LuminanceEffect lumeffect = new A.LuminanceEffect();
                        if (Picture.Brightness != 0) lumeffect.Brightness = Convert.ToInt32(Picture.Brightness * 1000);
                        if (Picture.Contrast != 0) lumeffect.Contrast = Convert.ToInt32(Picture.Contrast * 1000);
                        pic.BlipFill.Blip.Append(lumeffect);
                    }

                    pic.BlipFill.SourceRectangle = new A.SourceRectangle();
                    pic.BlipFill.Append(new A.Stretch() { FillRectangle = new A.FillRectangle() });

                    Picture.ShapeProperties.BlackWhiteMode = A.BlackWhiteModeValues.Auto;
                    Picture.ShapeProperties.HasTransform2D = true;
                    Picture.ShapeProperties.Transform2D.HasOffset = true;
                    Picture.ShapeProperties.Transform2D.HasExtents = true;

                    // not supporting yet because you need to change the positional offsets too.
                    //if (Picture.RotationAngle != 0)
                    //{
                    //    pic.ShapeProperties.Transform2D.Rotation = Convert.ToInt32(Picture.RotationAngle * (decimal)SLConstants.DegreeToAngleRepresentation);
                    //}

                    // used when it's relative positioning
                    // these are the actual values used, so it's 1 less than the given anchor indices.
                    int iColumnId = 0, iRowId = 0;
                    long lColumnOffset = 0, lRowOffset = 0;
                    if (Picture.UseRelativePositioning)
                    {
                        iColumnId = Picture.AnchorColumnIndex - 1;
                        iRowId = Picture.AnchorRowIndex - 1;

                        long lOffset = 0;
                        long lOffsetBuffer = 0;
                        int i;

                        for (i = 1; i <= iColumnId; ++i)
                        {
                            if (slws.ColumnProperties.ContainsKey(i))
                            {
                                cp = slws.ColumnProperties[i];
                                if (cp.HasWidth)
                                {
                                    lOffsetBuffer += cp.WidthInEMU;
                                }
                                else
                                {
                                    lOffsetBuffer += slws.SheetFormatProperties.DefaultColumnWidthInEMU;
                                }
                            }
                            else
                            {
                                // we use the current worksheet's column width
                                lOffsetBuffer += slws.SheetFormatProperties.DefaultColumnWidthInEMU;
                            }
                        }
                        lOffsetBuffer += Picture.OffsetX;
                        lOffset = lOffsetBuffer;

                        if (lOffset <= 0)
                        {
                            // in case the given offset is so negative, it pushes the sum to negative
                            // We use "<= 0" here, so the else part assumes a positive offset
                            iColumnId = 0;
                            lColumnOffset = 0;
                        }
                        else
                        {
                            lOffsetBuffer = 0;
                            i = 1;

                            while (lOffset > lOffsetBuffer)
                            {
                                iColumnId = i - 1;
                                lColumnOffset = lOffset - lOffsetBuffer;

                                if (slws.ColumnProperties.ContainsKey(i))
                                {
                                    cp = slws.ColumnProperties[i];
                                    if (cp.HasWidth)
                                    {
                                        lOffsetBuffer += cp.WidthInEMU;
                                    }
                                    else
                                    {
                                        lOffsetBuffer += slws.SheetFormatProperties.DefaultColumnWidthInEMU;
                                    }
                                }
                                else
                                {
                                    // we use the current worksheet's column width
                                    lOffsetBuffer += slws.SheetFormatProperties.DefaultColumnWidthInEMU;
                                }
                                ++i;
                            }
                        }

                        Picture.ShapeProperties.Transform2D.Offset.X = lColumnOffset;

                        lOffsetBuffer = 0;
                        for (i = 1; i <= iRowId; ++i)
                        {
                            if (slws.RowProperties.ContainsKey(i))
                            {
                                rp = slws.RowProperties[i];
                                if (rp.HasHeight)
                                {
                                    lOffsetBuffer += rp.HeightInEMU;
                                }
                                else
                                {
                                    lOffsetBuffer += slws.SheetFormatProperties.DefaultRowHeightInEMU;
                                }
                            }
                            else
                            {
                                // we use the current worksheet's row height
                                lOffsetBuffer += slws.SheetFormatProperties.DefaultRowHeightInEMU;
                            }
                        }
                        lOffsetBuffer += Picture.OffsetY;
                        lOffset = lOffsetBuffer;

                        if (lOffset <= 0)
                        {
                            // in case the given offset is so negative, it pushes the sum to negative
                            // We use "<= 0" here, so the else part assumes a positive offset
                            iRowId = 0;
                            lRowOffset = 0;
                        }
                        else
                        {
                            lOffsetBuffer = 0;
                            i = 1;

                            while (lOffset > lOffsetBuffer)
                            {
                                iRowId = i - 1;
                                lRowOffset = lOffset - lOffsetBuffer;

                                if (slws.RowProperties.ContainsKey(i))
                                {
                                    rp = slws.RowProperties[i];
                                    if (rp.HasHeight)
                                    {
                                        lOffsetBuffer += rp.HeightInEMU;
                                    }
                                    else
                                    {
                                        lOffsetBuffer += slws.SheetFormatProperties.DefaultRowHeightInEMU;
                                    }
                                }
                                else
                                {
                                    // we use the current worksheet's row height
                                    lOffsetBuffer += slws.SheetFormatProperties.DefaultRowHeightInEMU;
                                }
                                ++i;
                            }
                        }

                        Picture.ShapeProperties.Transform2D.Offset.Y = lRowOffset;
                    }
                    else
                    {
                        Picture.ShapeProperties.Transform2D.Offset.X = 0;
                        Picture.ShapeProperties.Transform2D.Offset.Y = 0;
                    }

                    Picture.ShapeProperties.Transform2D.Extents.Cx = Picture.WidthInEMU;
                    Picture.ShapeProperties.Transform2D.Extents.Cy = Picture.HeightInEMU;

                    Picture.ShapeProperties.HasPresetGeometry = true;
                    Picture.ShapeProperties.PresetGeometry = Picture.PictureShape;

                    pic.ShapeProperties = Picture.ShapeProperties.ToXdrShapeProperties();

                    Xdr.ClientData clientdata = new Xdr.ClientData();
                    // the properties are true by default
                    if (!Picture.LockWithSheet) clientdata.LockWithSheet = false;
                    if (!Picture.PrintWithSheet) clientdata.PrintWithSheet = false;

                    if (Picture.UseRelativePositioning)
                    {
                        Xdr.OneCellAnchor ocanchor = new Xdr.OneCellAnchor();
                        ocanchor.FromMarker = new Xdr.FromMarker();
                        // Subtract 1 because picture goes to bottom right corner
                        // Subtracting 1 makes it more intuitive that (1,1) means top-left corner of (1,1)
                        ocanchor.FromMarker.ColumnId = new Xdr.ColumnId() { Text = iColumnId.ToString(CultureInfo.InvariantCulture) };
                        ocanchor.FromMarker.ColumnOffset = new Xdr.ColumnOffset() { Text = lColumnOffset.ToString(CultureInfo.InvariantCulture) };
                        ocanchor.FromMarker.RowId = new Xdr.RowId() { Text = iRowId.ToString(CultureInfo.InvariantCulture) };
                        ocanchor.FromMarker.RowOffset = new Xdr.RowOffset() { Text = lRowOffset.ToString(CultureInfo.InvariantCulture) };

                        ocanchor.Extent = new Xdr.Extent();
                        ocanchor.Extent.Cx = Picture.WidthInEMU;
                        ocanchor.Extent.Cy = Picture.HeightInEMU;

                        ocanchor.Append(pic);
                        ocanchor.Append(clientdata);
                        wsd.Append(ocanchor);
                    }
                    else
                    {
                        Xdr.AbsoluteAnchor absanchor = new Xdr.AbsoluteAnchor();
                        absanchor.Position = new Xdr.Position();
                        absanchor.Position.X = Picture.OffsetX;
                        absanchor.Position.Y = Picture.OffsetY;

                        absanchor.Extent = new Xdr.Extent();
                        absanchor.Extent.Cx = Picture.WidthInEMU;
                        absanchor.Extent.Cy = Picture.HeightInEMU;

                        absanchor.Append(pic);
                        absanchor.Append(clientdata);
                        wsd.Append(absanchor);
                    }

                    wsd.Save(dp);
                }
            }
            #endregion
        }

        /// <summary>
        /// Adds conditional formatting into the currently selected worksheet.
        /// </summary>
        /// <param name="ConditionalFormatting">An SLConditionalFormatting object with the formatting rules already set. Remember to set at least one formatting rule (a data bar, color scale, icon set or some custom rule).</param>
        /// <returns>True if successfully added. False otherwise.</returns>
        public bool AddConditionalFormatting(SLConditionalFormatting ConditionalFormatting)
        {
            bool result = false;
            if (ConditionalFormatting.Rules.Count > 0 && ConditionalFormatting.SequenceOfReferences.Count > 0)
            {
                result = true;
                int index = ConditionalFormatting.Rules.Count;
                foreach (SLConditionalFormatting cf in slws.ConditionalFormattings)
                {
                    foreach (SLConditionalFormattingRule cfr in cf.Rules)
                    {
                        cfr.Priority += index;
                    }
                }

                foreach (SLConditionalFormatting2010 cf2010 in slws.ConditionalFormattings2010)
                {
                    foreach (SLConditionalFormattingRule2010 cfr2010 in cf2010.Rules)
                    {
                        if (cfr2010.Priority != null) cfr2010.Priority += index;
                    }
                }

                bool bIs2010 = false;
                bool bIsDataBar2010 = false;
                bool bIsIconSet2010 = false;
                string sGuid = string.Empty;
                SLConditionalFormatting2010 cf2010new = new SLConditionalFormatting2010();
                foreach (SLCellPointRange pt in ConditionalFormatting.SequenceOfReferences)
                {
                    cf2010new.ReferenceSequence.Add(new SLCellPointRange(pt.StartRowIndex, pt.StartColumnIndex, pt.EndRowIndex, pt.EndColumnIndex));
                }
                SLConditionalFormattingRule2010 cfr2010new;

                ConditionalFormattingRuleExtension cfrext;

                index = 1;
                int i;
                // the latest added rule takes first priority
                // And also we might delete (2010 icon set) so need to start at the end
                for (i = ConditionalFormatting.Rules.Count - 1; i >= 0; --i)
                {
                    ConditionalFormatting.Rules[i].Priority = index;
                    ++index;
                    if (ConditionalFormatting.Rules[i].HasDifferentialFormat)
                    {
                        ConditionalFormatting.Rules[i].FormatId = (uint)this.SaveToStylesheetDifferentialFormat(ConditionalFormatting.Rules[i].DifferentialFormat.ToHash());
                    }

                    bIsDataBar2010 = ConditionalFormatting.Rules[i].HasDataBar && ConditionalFormatting.Rules[i].DataBar.Is2010;
                    bIsIconSet2010 = ConditionalFormatting.Rules[i].HasIconSet && ConditionalFormatting.Rules[i].IconSet.Is2010;

                    // supposedly both cannot be true at the same time because each rule
                    // can only have one rule (duh) at any one time, whether it's color scale,
                    // data bar, icon set, top 10 or whatever rule there is.
                    if (bIsDataBar2010 || bIsIconSet2010)
                    {
                        bIs2010 = true;
                        cfr2010new = ConditionalFormatting.Rules[i].ToSLConditionalFormattingRule2010();

                        if (bIsDataBar2010)
                        {
                            // go read the Open XML specs on why it has to be null.
                            // We null the priority so the extension fill color is not rendered.
                            // Presumably, it uses the color in the normally placed data bar.
                            // I don't know why Microsoft made it such that data bars exist in the
                            // normal place *and* in the extension place.
                            // Why not exist entirely in either the normal or extension like the
                            // Excel 2010 icon sets?
                            cfr2010new.Priority = null;

                            // If I'm reading the specs correctly, if there's an extension, then
                            // the normal placed data bar needs to have the min/max lengths as defaults.
                            // AKA 10 and 90 percent respectively.
                            // Hey reading that part of the specs takes longer than watching the Titanic
                            // movie... and more convoluted than a legal clause...
                            ConditionalFormatting.Rules[i].DataBar.MinLength = 10;
                            ConditionalFormatting.Rules[i].DataBar.MaxLength = 90;

                            sGuid = string.Format("{{{0}}}", Guid.NewGuid().ToString().ToUpperInvariant());
                            cfr2010new.Id = sGuid;

                            cfrext = new ConditionalFormattingRuleExtension();
                            cfrext.Uri = SLConstants.DataBarExtensionUri;
                            cfrext.Append(new X14.Id(sGuid));
                            ConditionalFormatting.Rules[i].Extensions.Add(cfrext);
                        }

                        if (bIsIconSet2010)
                        {
                            sGuid = string.Format("{{{0}}}", Guid.NewGuid().ToString().ToUpperInvariant());
                            cfr2010new.Id = sGuid;
                            // 2010 icon set exists entirely in the extension part
                            ConditionalFormatting.Rules.RemoveAt(i);
                        }

                        // we insert at index 0 because we started from the end of the existing set
                        // of rules. The new 2010 version will then follow the existing order (sort of).
                        cf2010new.Rules.Insert(0, cfr2010new.Clone());
                    }
                }

                // in case there's only 2010 icon set rule
                if (ConditionalFormatting.Rules.Count > 0)
                {
                    slws.ConditionalFormattings.Add(ConditionalFormatting.Clone());
                }

                if (bIs2010)
                {
                    slws.ConditionalFormattings2010.Add(cf2010new.Clone());
                }
            }

            return result;
        }

        /// <summary>
        /// Clear all conditional formatting from the currently selected worksheet.
        /// </summary>
        public void ClearConditionalFormatting()
        {
            slws.ConditionalFormattings.Clear();
            slws.ConditionalFormattings2010.Clear();
        }

        /// <summary>
        /// Adds data validation into the currently selected worksheet.
        /// </summary>
        /// <param name="DataValidation">An SLDataValidation with desired settings.</param>
        /// <returns>True if successful. False otherwise. Failure is probably due to overlapping data validation regions.</returns>
        public bool AddDataValidation(SLDataValidation DataValidation)
        {
            List<SLCellPointRange> listExisting = new List<SLCellPointRange>();
            foreach (SLDataValidation dv in slws.DataValidations)
            {
                listExisting.AddRange(dv.SequenceOfReferences);
            }

            // not checking against itself on list of cell references.
            // Just check against other existing data validations.
            // As of this writing, there's no way to add in extra cell references anyway...
            SLCellPointRange pt = DataValidation.SequenceOfReferences[0];

            // This is the separating axis theorem. See merging cells for more details.
            // For simplicity, we don't allow overlapping. Excel seems to just override existing
            // data validations with the new range.
            int i;
            bool result = true;
            for (i = 0; i < listExisting.Count; ++i)
            {
                if (!(pt.EndRowIndex < listExisting[i].StartRowIndex || pt.StartRowIndex > listExisting[i].EndRowIndex
                    || pt.EndColumnIndex < listExisting[i].StartColumnIndex || pt.StartColumnIndex > listExisting[i].EndColumnIndex))
                {
                    result = false;
                    break;
                }
            }

            if (result)
            {
                slws.DataValidations.Add(DataValidation.Clone());
            }

            return result;
        }

        /// <summary>
        /// Clear all data validations from the currently selected worksheet.
        /// </summary>
        public void ClearDataValidation()
        {
            slws.DataValidations.Clear();
        }

        /// <summary>
        /// Insert a table into the currently selected worksheet.
        /// </summary>
        /// <param name="Table">An SLTable object with the properties already set.</param>
        /// <returns>True if the table is successfully inserted. False otherwise. If it failed, check if the given table overlaps any existing tables or merged cell range.</returns>
        public bool InsertTable(SLTable Table)
        {
            // This is the separating axis theorem. See merging cells for more details.
            // We're checking if the table collides with merged cells, the worksheet's autofilter range
            // and existing tables.
            // Technically, Excel unmerges cells when a table overlaps a merged cell range.
            // We're just going to fail that.

            bool result = true;
            int i, j;
            for (i = 0; i < slws.MergeCells.Count; ++i)
            {
                if (!(Table.EndRowIndex < slws.MergeCells[i].StartRowIndex
                    || Table.StartRowIndex > slws.MergeCells[i].EndRowIndex
                    || Table.EndColumnIndex < slws.MergeCells[i].StartColumnIndex
                    || Table.StartColumnIndex > slws.MergeCells[i].EndColumnIndex))
                {
                    result = false;
                    break;
                }
            }

            if (slws.HasAutoFilter)
            {
                if (!(Table.EndRowIndex < slws.AutoFilter.StartRowIndex
                    || Table.StartRowIndex > slws.AutoFilter.EndRowIndex
                    || Table.EndColumnIndex < slws.AutoFilter.StartColumnIndex
                    || Table.StartColumnIndex > slws.AutoFilter.EndColumnIndex))
                {
                    result = false;
                }
            }

            if (!result) return false;

            for (i = 0; i < slws.Tables.Count; ++i)
            {
                if (!(Table.EndRowIndex < slws.Tables[i].StartRowIndex
                    || Table.StartRowIndex > slws.Tables[i].EndRowIndex
                    || Table.EndColumnIndex < slws.Tables[i].StartColumnIndex
                    || Table.StartColumnIndex > slws.Tables[i].EndColumnIndex))
                {
                    result = false;
                    break;
                }
            }

            if (result)
            {
                // sorting first!
                // We'll do just one level deep sorting. Multiple level sorting is hard...
                if (Table.HasSortState && Table.SortState.SortConditions.Count > 0)
                {
                    bool bSortAscending = true;
                    if (Table.SortState.SortConditions[0].Descending != null)
                        bSortAscending = !Table.SortState.SortConditions[0].Descending.Value;

                    this.Sort(Table.SortState.StartRowIndex, Table.SortState.StartColumnIndex,
                        Table.SortState.EndRowIndex, Table.SortState.EndColumnIndex, true,
                        Table.SortState.SortConditions[0].StartColumnIndex, bSortAscending);
                }

                // filtering next! Because rows might be hidden

                int iStartRowIndex = -1;
                int iEndRowIndex = -1;
                if (Table.HeaderRowCount > 0) iStartRowIndex = Table.StartRowIndex + 1;
                else iStartRowIndex = Table.StartRowIndex;
                // not inclusive of the last totals row
                iEndRowIndex = Table.EndRowIndex - 1;

                SLTableColumn tc;
                List<SLCell> cells;
                SLCell c;
                string sResultText = string.Empty;
                SLCalculationCell cc;

                uint iStyleIndex;

                for (j = 0; j < Table.TableColumns.Count; ++j)
                {
                    tc = Table.TableColumns[j];
                    if (tc.TotalsRowLabel != null && tc.TotalsRowLabel.Length > 0)
                    {
                        c = new SLCell();
                        c.DataType = CellValues.SharedString;
                        c.NumericValue = this.DirectSaveToSharedStringTable(SLTool.XmlWrite(tc.TotalsRowLabel, gbThrowExceptionsIfAny));
                        slws.CellWarehouse.SetValue(Table.EndRowIndex, Table.StartColumnIndex + j, c);
                    }
                    if (tc.HasTotalsRowFunction)
                    {
                        cells = new List<SLCell>();
                        for (i = iStartRowIndex; i <= iEndRowIndex; ++i)
                        {
                            if (slws.CellWarehouse.Exists(i, Table.StartColumnIndex + j))
                            {
                                cells.Add(slws.CellWarehouse.Cells[i][Table.StartColumnIndex + j].Clone());
                            }
                        }

                        c = new SLCell();
                        c.CellFormula = new SLCellFormula();
                        c.CellFormula.FormulaText = string.Format("SUBTOTAL({0},[{1}])", this.GetFunctionNumber(tc.TotalsRowFunction), tc.Name);
                        if (!this.Calculate(tc.TotalsRowFunction, cells, out sResultText))
                        {
                            c.DataType = CellValues.Error;
                        }
                        c.CellText = sResultText;

                        iStyleIndex = 0;
                        if (slws.RowProperties.ContainsKey(Table.EndRowIndex)) iStyleIndex = slws.RowProperties[Table.EndRowIndex].StyleIndex;
                        if (iStyleIndex == 0 && slws.ColumnProperties.ContainsKey(Table.StartColumnIndex + j)) iStyleIndex = slws.ColumnProperties[Table.StartColumnIndex + j].StyleIndex;
                        if (iStyleIndex != 0) c.StyleIndex = (uint)iStyleIndex;

                        slws.CellWarehouse.SetValue(Table.EndRowIndex, Table.StartColumnIndex + j, c);

                        cc = new SLCalculationCell(SLTool.ToCellReference(Table.EndRowIndex, Table.StartColumnIndex + j));
                        cc.SheetId = (int)giSelectedWorksheetID;
                        slwb.AddCalculationCell(cc);
                    }
                }

                if (slwb.HasTableName(Table.DisplayName) || Table.DisplayName.Contains(" "))
                {
                    slwb.RefreshPossibleTableId();
                    Table.Id = slwb.PossibleTableId;
                    Table.sDisplayName = string.Format("Table{0}", Table.Id.ToString(CultureInfo.InvariantCulture));
                    Table.Name = Table.sDisplayName;
                }

                if (!slwb.TableIds.Contains(Table.Id)) slwb.TableIds.Add(Table.Id);

                if (!slwb.TableNames.Contains(Table.DisplayName)) slwb.TableNames.Add(Table.DisplayName);

                slws.Tables.Add(Table.Clone());
            }

            return result;
        }

        /// <summary>
        /// Get the page settings of the currently selected worksheet.
        /// </summary>
        /// <returns>An SLPageSettings object with the page settings of the currently selected worksheet.</returns>
        public SLPageSettings GetPageSettings()
        {
            return slws.PageSettings.Clone();
        }

        /// <summary>
        /// Get the page settings of sheet.
        /// </summary>
        /// <param name="SheetName">The name of the sheet.</param>
        /// <returns>An SLPageSettings object with the page settings of the specified sheet.</returns>
        public SLPageSettings GetPageSettings(string SheetName)
        {
            if (SheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                return slws.PageSettings.Clone();
            }

            SLPageSettings ps = new SLPageSettings(SimpleTheme.listThemeColors, SimpleTheme.listIndexedColors);

            bool bSheetFound = false;
            SLSheet sheet = new SLSheet(string.Empty, 0, string.Empty, SLSheetType.Unknown);
            foreach (SLSheet s in slwb.Sheets)
            {
                if (s.Name.Equals(SheetName, StringComparison.OrdinalIgnoreCase))
                {
                    bSheetFound = true;
                    sheet = s.Clone();
                    break;
                }
            }

            if (bSheetFound)
            {
                if (sheet.SheetType == SLSheetType.Worksheet)
                {
                    WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(sheet.Id);
                    using (OpenXmlReader oxr = OpenXmlReader.Create(wsp))
                    {
                        bool bFound = false;
                        SLSheetView slsv;
                        while (oxr.Read())
                        {
                            if (oxr.ElementType == typeof(SheetProperties))
                            {
                                ps.SheetProperties.FromSheetProperties((SheetProperties)oxr.LoadCurrentElement());
                            }
                            else if (oxr.ElementType == typeof(SheetView))
                            {
                                // if we find a sheet view with the default workbook view id,
                                // we'll just take that.
                                if (!bFound)
                                {
                                    slsv = new SLSheetView();
                                    slsv.FromSheetView((SheetView)oxr.LoadCurrentElement());
                                    if (slsv.ShowFormulas) ps.bShowFormulas = slsv.ShowFormulas;
                                    if (!slsv.ShowGridLines) ps.bShowGridLines = slsv.ShowGridLines;
                                    if (!slsv.ShowRowColHeaders) ps.bShowRowColumnHeaders = slsv.ShowRowColHeaders;
                                    if (!slsv.ShowRuler) ps.bShowRuler = slsv.ShowRuler;
                                    if (slsv.View != SheetViewValues.Normal) ps.vView = slsv.View;
                                    if (slsv.ZoomScale != 100) ps.ZoomScale = slsv.ZoomScale;
                                    if (slsv.ZoomScaleNormal != 0) ps.ZoomScaleNormal = slsv.ZoomScaleNormal;
                                    if (slsv.ZoomScalePageLayoutView != 0) ps.ZoomScalePageLayoutView = slsv.ZoomScalePageLayoutView;

                                    if (slsv.WorkbookViewId == 0) bFound = true;
                                }
                            }
                            else if (oxr.ElementType == typeof(PrintOptions))
                            {
                                ps.ImportPrintOptions((PrintOptions)oxr.LoadCurrentElement());
                            }
                            else if (oxr.ElementType == typeof(PageMargins))
                            {
                                ps.ImportPageMargins((PageMargins)oxr.LoadCurrentElement());
                            }
                            else if (oxr.ElementType == typeof(PageSetup))
                            {
                                ps.ImportPageSetup((PageSetup)oxr.LoadCurrentElement());
                            }
                            else if (oxr.ElementType == typeof(HeaderFooter))
                            {
                                ps.ImportHeaderFooter((HeaderFooter)oxr.LoadCurrentElement());
                            }
                        }
                    }
                }
                else if (sheet.SheetType == SLSheetType.Chartsheet)
                {
                    ChartsheetPart csp = (ChartsheetPart)wbp.GetPartById(sheet.Id);
                    if (csp.Chartsheet.ChartSheetProperties != null)
                    {
                        ps.SheetProperties.FromChartSheetProperties(csp.Chartsheet.ChartSheetProperties);
                    }
                    if (csp.Chartsheet.PageMargins != null)
                    {
                        ps.ImportPageMargins(csp.Chartsheet.PageMargins);
                    }
                    if (csp.Chartsheet.ChartSheetPageSetup != null)
                    {
                        ps.ImportChartSheetPageSetup(csp.Chartsheet.ChartSheetPageSetup);
                    }
                    if (csp.Chartsheet.HeaderFooter != null)
                    {
                        ps.ImportHeaderFooter(csp.Chartsheet.HeaderFooter);
                    }
                }
                else if (sheet.SheetType == SLSheetType.DialogSheet)
                {
                    DialogsheetPart dsp = (DialogsheetPart)wbp.GetPartById(sheet.Id);
                    if (dsp.DialogSheet.SheetProperties != null)
                    {
                        ps.SheetProperties.FromSheetProperties(dsp.DialogSheet.SheetProperties);
                    }
                    if (dsp.DialogSheet.PrintOptions != null)
                    {
                        ps.ImportPrintOptions(dsp.DialogSheet.PrintOptions);
                    }
                    if (dsp.DialogSheet.PageMargins != null)
                    {
                        ps.ImportPageMargins(dsp.DialogSheet.PageMargins);
                    }
                    if (dsp.DialogSheet.PageSetup != null)
                    {
                        ps.ImportPageSetup(dsp.DialogSheet.PageSetup);
                    }
                    if (dsp.DialogSheet.HeaderFooter != null)
                    {
                        ps.ImportHeaderFooter(dsp.DialogSheet.HeaderFooter);
                    }
                }
                else if (sheet.SheetType == SLSheetType.Macrosheet)
                {
                    // not doing anything for macrosheets. What *are* macrosheets?
                }
            }

            return ps;
        }

        internal void SetPageSettingsSheetView(SLPageSettings ps)
        {
            if (ps.bShowFormulas != null)
            {
                // TODO: images and charts?
                // Actually I don't feel like updating those...

                if (ps.bShowFormulas.Value != slws.IsDoubleColumnWidth)
                {
                    List<int> keys = slws.ColumnProperties.Keys.ToList<int>();
                    SLColumnProperties cp;
                    slws.IsDoubleColumnWidth = ps.bShowFormulas.Value;
                    if (ps.bShowFormulas.Value)
                    {
                        // have to test beforehand because setting the default column width
                        // assigns the HasDefaultColumnWidth property
                        if (!slws.SheetFormatProperties.HasDefaultColumnWidth)
                        {
                            slws.SheetFormatProperties.DefaultColumnWidth = 2 * slws.SheetFormatProperties.DefaultColumnWidth;
                            slws.SheetFormatProperties.HasDefaultColumnWidth = false;
                        }
                        else
                        {
                            slws.SheetFormatProperties.DefaultColumnWidth = 2 * slws.SheetFormatProperties.DefaultColumnWidth;
                        }

                        foreach (int colkey in keys)
                        {
                            cp = slws.ColumnProperties[colkey];
                            if (cp.HasWidth)
                            {
                                cp.Width = 2 * cp.Width;
                                slws.ColumnProperties[colkey] = cp.Clone();
                            }
                        }
                    }
                    else
                    {
                        // have to test beforehand because setting the default column width
                        // assigns the HasDefaultColumnWidth property
                        if (!slws.SheetFormatProperties.HasDefaultColumnWidth)
                        {
                            slws.SheetFormatProperties.DefaultColumnWidth = 0.5 * slws.SheetFormatProperties.DefaultColumnWidth;
                            slws.SheetFormatProperties.HasDefaultColumnWidth = false;
                        }
                        else
                        {
                            slws.SheetFormatProperties.DefaultColumnWidth = 0.5 * slws.SheetFormatProperties.DefaultColumnWidth;
                        }

                        foreach (int colkey in keys)
                        {
                            cp = slws.ColumnProperties[colkey];
                            if (cp.HasWidth)
                            {
                                cp.Width = 0.5 * cp.Width;
                                slws.ColumnProperties[colkey] = cp.Clone();
                            }
                        }
                    }
                }
            }

            if (ps.HasSheetView)
            {
                if (slws.SheetViews.Count > 0)
                {
                    bool bFound = false;
                    foreach (SLSheetView sv in slws.SheetViews)
                    {
                        if (sv.WorkbookViewId == 0)
                        {
                            bFound = true;
                            if (ps.bShowFormulas != null) sv.ShowFormulas = ps.bShowFormulas.Value;
                            if (ps.bShowGridLines != null) sv.ShowGridLines = ps.bShowGridLines.Value;
                            if (ps.bShowRowColumnHeaders != null) sv.ShowRowColHeaders = ps.bShowRowColumnHeaders.Value;
                            if (ps.bShowRuler != null) sv.ShowRuler = ps.bShowRuler.Value;
                            if (ps.vView != null) sv.View = ps.vView.Value;
                            if (ps.iZoomScale != null) sv.ZoomScale = ps.iZoomScale.Value;
                            if (ps.iZoomScaleNormal != null) sv.ZoomScaleNormal = ps.iZoomScaleNormal.Value;
                            if (ps.iZoomScalePageLayoutView != null) sv.ZoomScalePageLayoutView = ps.iZoomScalePageLayoutView.Value;
                        }
                    }

                    if (!bFound)
                    {
                        slws.SheetViews.Add(ps.ExportSLSheetView());
                    }
                }
                else
                {
                    slws.SheetViews.Add(ps.ExportSLSheetView());
                }
            }
        }

        /// <summary>
        /// Set page settings to the currently selected worksheet.
        /// </summary>
        /// <param name="PageSettings">An SLPageSettings object with the properties already set.</param>
        public void SetPageSettings(SLPageSettings PageSettings)
        {
            slws.PageSettings = PageSettings.Clone();
            this.SetPageSettingsSheetView(PageSettings);
        }

        /// <summary>
        /// Set page settings to a sheet.
        /// </summary>
        /// <param name="PageSettings">An SLPageSettings object with the properties already set.</param>
        /// <param name="SheetName">The name of the sheet.</param>
        public void SetPageSettings(SLPageSettings PageSettings, string SheetName)
        {
            if (SheetName.Equals(gsSelectedWorksheetName, StringComparison.OrdinalIgnoreCase))
            {
                slws.PageSettings = PageSettings.Clone();
                this.SetPageSettingsSheetView(PageSettings);
                return;
            }

            // we're not going to double column widths for non-currently-selected worksheets
            // when show formulas is true...
            // Too much work for the rare occurence that it happens...

            bool bSheetFound = false;
            SLSheet sheet = new SLSheet(string.Empty, 0, string.Empty, SLSheetType.Unknown);
            foreach (SLSheet s in slwb.Sheets)
            {
                if (s.Name.Equals(SheetName, StringComparison.OrdinalIgnoreCase))
                {
                    bSheetFound = true;
                    sheet = s.Clone();
                    break;
                }
            }

            if (bSheetFound)
            {
                bool bFound = false;
                OpenXmlElement oxe;

                if (sheet.SheetType == SLSheetType.Worksheet)
                {
                    #region Worksheet
                    WorksheetPart wsp = (WorksheetPart)wbp.GetPartById(sheet.Id);
                    if (PageSettings.HasSheetProperties)
                    {
                        wsp.Worksheet.SheetProperties = PageSettings.SheetProperties.ToSheetProperties();
                    }
                    else
                    {
                        wsp.Worksheet.SheetProperties = null;
                    }

                    #region SheetViews
                    if (PageSettings.HasSheetView)
                    {
                        if (wsp.Worksheet.SheetViews != null)
                        {
                            bool bSheetViewFound = false;
                            foreach (SheetView sv in wsp.Worksheet.SheetViews)
                            {
                                if (sv.WorkbookViewId == 0)
                                {
                                    if (PageSettings.bShowFormulas != null) sv.ShowFormulas = PageSettings.bShowFormulas.Value;
                                    if (PageSettings.bShowGridLines != null) sv.ShowGridLines = PageSettings.bShowGridLines.Value;
                                    if (PageSettings.bShowRowColumnHeaders != null) sv.ShowRowColHeaders = PageSettings.bShowRowColumnHeaders.Value;
                                    if (PageSettings.bShowRuler != null) sv.ShowRuler = PageSettings.bShowRuler.Value;
                                    if (PageSettings.vView != null) sv.View = PageSettings.vView.Value;
                                    if (PageSettings.iZoomScale != null) sv.ZoomScale = PageSettings.iZoomScale.Value;
                                    if (PageSettings.iZoomScaleNormal != null) sv.ZoomScaleNormal = PageSettings.iZoomScaleNormal.Value;
                                    if (PageSettings.iZoomScalePageLayoutView != null) sv.ZoomScalePageLayoutView = PageSettings.iZoomScalePageLayoutView.Value;
                                }
                            }

                            if (!bSheetViewFound)
                            {
                                wsp.Worksheet.SheetViews.Append(PageSettings.ExportSLSheetView().ToSheetView());
                            }
                        }
                        else
                        {
                            wsp.Worksheet.SheetViews = new SheetViews();
                            wsp.Worksheet.SheetViews.Append(PageSettings.ExportSLSheetView().ToSheetView());
                        }
                    }
                    #endregion

                    #region PrintOptions
                    if (PageSettings.HasPrintOptions)
                    {
                        if (wsp.Worksheet.Elements<PrintOptions>().Count() > 0)
                        {
                            wsp.Worksheet.RemoveAllChildren<PrintOptions>();
                        }

                        bFound = false;
                        oxe = wsp.Worksheet.FirstChild;
                        foreach (var child in wsp.Worksheet.ChildElements)
                        {
                            // start with SheetData because it's a required child element
                            if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                                || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                                || child is SortState || child is DataConsolidate || child is CustomSheetViews
                                || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                                || child is DataValidations || child is Hyperlinks)
                            {
                                oxe = child;
                                bFound = true;
                            }
                        }

                        if (bFound)
                        {
                            wsp.Worksheet.InsertAfter(PageSettings.ExportPrintOptions(), oxe);
                        }
                        else
                        {
                            wsp.Worksheet.PrependChild(PageSettings.ExportPrintOptions());
                        }
                    }
                    #endregion

                    #region PageMargins
                    if (PageSettings.HasPageMargins)
                    {
                        if (wsp.Worksheet.Elements<PageMargins>().Count() > 0)
                        {
                            wsp.Worksheet.RemoveAllChildren<PageMargins>();
                        }

                        bFound = false;
                        oxe = wsp.Worksheet.FirstChild;
                        foreach (var child in wsp.Worksheet.ChildElements)
                        {
                            // start with SheetData because it's a required child element
                            if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                                || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                                || child is SortState || child is DataConsolidate || child is CustomSheetViews
                                || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                                || child is DataValidations || child is Hyperlinks || child is PrintOptions)
                            {
                                oxe = child;
                                bFound = true;
                            }
                        }

                        if (bFound)
                        {
                            wsp.Worksheet.InsertAfter(PageSettings.ExportPageMargins(), oxe);
                        }
                        else
                        {
                            wsp.Worksheet.PrependChild(PageSettings.ExportPageMargins());
                        }
                    }
                    #endregion

                    #region PageSetup
                    if (PageSettings.HasPageSetup)
                    {
                        if (wsp.Worksheet.Elements<PageSetup>().Count() > 0)
                        {
                            wsp.Worksheet.RemoveAllChildren<PageSetup>();
                        }

                        bFound = false;
                        oxe = wsp.Worksheet.FirstChild;
                        foreach (var child in wsp.Worksheet.ChildElements)
                        {
                            // start with SheetData because it's a required child element
                            if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                                || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                                || child is SortState || child is DataConsolidate || child is CustomSheetViews
                                || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                                || child is DataValidations || child is Hyperlinks || child is PrintOptions
                                || child is PageMargins)
                            {
                                oxe = child;
                                bFound = true;
                            }
                        }

                        if (bFound)
                        {
                            wsp.Worksheet.InsertAfter(PageSettings.ExportPageSetup(), oxe);
                        }
                        else
                        {
                            wsp.Worksheet.PrependChild(PageSettings.ExportPageSetup());
                        }
                    }
                    #endregion

                    #region HeaderFooter
                    if (PageSettings.HasHeaderFooter)
                    {
                        if (wsp.Worksheet.Elements<HeaderFooter>().Count() > 0)
                        {
                            wsp.Worksheet.RemoveAllChildren<HeaderFooter>();
                        }

                        bFound = false;
                        oxe = wsp.Worksheet.FirstChild;
                        foreach (var child in wsp.Worksheet.ChildElements)
                        {
                            // start with SheetData because it's a required child element
                            if (child is SheetData || child is SheetCalculationProperties || child is SheetProtection
                                || child is ProtectedRanges || child is Scenarios || child is AutoFilter
                                || child is SortState || child is DataConsolidate || child is CustomSheetViews
                                || child is MergeCells || child is PhoneticProperties || child is ConditionalFormatting
                                || child is DataValidations || child is Hyperlinks || child is PrintOptions
                                || child is PageMargins || child is PageSetup)
                            {
                                oxe = child;
                                bFound = true;
                            }
                        }

                        if (bFound)
                        {
                            wsp.Worksheet.InsertAfter(PageSettings.ExportHeaderFooter(), oxe);
                        }
                        else
                        {
                            wsp.Worksheet.PrependChild(PageSettings.ExportHeaderFooter());
                        }
                    }
                    #endregion

                    wsp.Worksheet.Save();
                    #endregion
                }
                else if (sheet.SheetType == SLSheetType.Chartsheet)
                {
                    #region Chartsheet
                    ChartsheetPart csp = (ChartsheetPart)wbp.GetPartById(sheet.Id);
                    if (PageSettings.HasChartSheetProperties)
                    {
                        csp.Chartsheet.ChartSheetProperties = PageSettings.SheetProperties.ToChartSheetProperties();
                    }
                    else
                    {
                        csp.Chartsheet.ChartSheetProperties = null;
                    }

                    if (PageSettings.HasPageMargins)
                    {
                        csp.Chartsheet.PageMargins = PageSettings.ExportPageMargins();
                    }
                    else
                    {
                        csp.Chartsheet.PageMargins = null;
                    }

                    if (PageSettings.HasChartSheetPageSetup)
                    {
                        csp.Chartsheet.ChartSheetPageSetup = PageSettings.ExportChartSheetPageSetup();
                    }
                    else
                    {
                        csp.Chartsheet.ChartSheetPageSetup = null;
                    }

                    if (PageSettings.HasHeaderFooter)
                    {
                        csp.Chartsheet.HeaderFooter = PageSettings.ExportHeaderFooter();
                    }
                    else
                    {
                        csp.Chartsheet.HeaderFooter = null;
                    }

                    csp.Chartsheet.Save();
                    #endregion
                }
                else if (sheet.SheetType == SLSheetType.DialogSheet)
                {
                    #region DialogSheet
                    DialogsheetPart dsp = (DialogsheetPart)wbp.GetPartById(sheet.Id);
                    if (PageSettings.HasSheetProperties)
                    {
                        dsp.DialogSheet.SheetProperties = PageSettings.SheetProperties.ToSheetProperties();
                    }
                    else
                    {
                        dsp.DialogSheet.SheetProperties = null;
                    }

                    if (PageSettings.HasPrintOptions)
                    {
                        dsp.DialogSheet.PrintOptions = PageSettings.ExportPrintOptions();
                    }
                    else
                    {
                        dsp.DialogSheet.PrintOptions = null;
                    }

                    if (PageSettings.HasPageMargins)
                    {
                        dsp.DialogSheet.PageMargins = PageSettings.ExportPageMargins();
                    }
                    else
                    {
                        dsp.DialogSheet.PageMargins = null;
                    }

                    if (PageSettings.HasPageSetup)
                    {
                        dsp.DialogSheet.PageSetup = PageSettings.ExportPageSetup();
                    }
                    else
                    {
                        dsp.DialogSheet.PageSetup = null;
                    }

                    if (PageSettings.HasHeaderFooter)
                    {
                        dsp.DialogSheet.HeaderFooter = PageSettings.ExportHeaderFooter();
                    }
                    else
                    {
                        dsp.DialogSheet.HeaderFooter = null;
                    }

                    dsp.DialogSheet.Save();
                    #endregion
                }
                else if (sheet.SheetType == SLSheetType.Macrosheet)
                {
                    #region Macrosheet
                    // don't care about macrosheets for now. What *are* macrosheets?
                    #endregion
                }
            }
        }

        /// <summary>
        /// Insert a page break above a given row index and to the left of a given column index.
        /// </summary>
        /// <param name="RowIndex">The row index. Use a negative value to ignore row breaks (-1 works fine).</param>
        /// <param name="ColumnIndex">The column index. Use a negative value to ignore column breaks (-1 works fine).</param>
        public void InsertPageBreak(int RowIndex, int ColumnIndex)
        {
            if (RowIndex < 1 || RowIndex > SLConstants.RowLimit) RowIndex = -1;
            if (ColumnIndex < 1 || ColumnIndex > SLConstants.ColumnLimit) ColumnIndex = -1;

            int iRowIndex = RowIndex - 1;
            int iColumnIndex = ColumnIndex - 1;

            // if both indices are out of range, just return.
            // This includes the case when A1 (row 1, column 1) is given too.
            if (iRowIndex <= 0 && iColumnIndex <= 0) return;

            SLBreak b;
            if (iRowIndex > 0)
            {
                b = new SLBreak();
                b.Id = (uint)iRowIndex;
                b.Max = (uint)(SLConstants.RowLimit - 1);
                b.ManualPageBreak = true;
                slws.RowBreaks[iRowIndex] = b;
            }

            if (iColumnIndex > 0)
            {
                b = new SLBreak();
                b.Id = (uint)iColumnIndex;
                b.Max = (uint)(SLConstants.ColumnLimit - 1);
                b.ManualPageBreak = true;
                slws.ColumnBreaks[iColumnIndex] = b;
            }
        }

        /// <summary>
        /// Remove all page breaks from the currently selected worksheet.
        /// </summary>
        public void RemoveAllPageBreaks()
        {
            slws.RowBreaks.Clear();
            slws.ColumnBreaks.Clear();
        }

        /// <summary>
        /// Remove a page break above a given row index and to the left of a given column index.
        /// </summary>
        /// <param name="RowIndex">The row index. Use a negative value to ignore row breaks (-1 works fine).</param>
        /// <param name="ColumnIndex">The column index. Use a negative value to ignore column breaks (-1 works fine).</param>
        public void RemovePageBreak(int RowIndex, int ColumnIndex)
        {
            int iRowIndex = RowIndex - 1;
            int iColumnIndex = ColumnIndex - 1;

            if (slws.RowBreaks.ContainsKey(iRowIndex))
            {
                slws.RowBreaks.Remove(iRowIndex);
            }

            if (slws.ColumnBreaks.ContainsKey(iColumnIndex))
            {
                slws.ColumnBreaks.Remove(iColumnIndex);
            }
        }

        /// <summary>
        /// Protect the currently selected worksheet. If the worksheet has protection (but not password protected), the current protection options will be overwritten.
        /// </summary>
        /// <param name="ProtectOptions">An SLSheetProtection object with relevant options set.</param>
        /// <returns>True if operation is successful. False otherwise. Note that if the worksheet already has password protection, false is also returned.</returns>
        public bool ProtectWorksheet(SLSheetProtection ProtectOptions)
        {
            bool result = true;
            if (slws.HasSheetProtection)
            {
                if (slws.SheetProtection.HashValue != null || slws.SheetProtection.Password != null)
                {
                    result = false;
                }
            }

            if (result)
            {
                ProtectOptions.Sheet = true;
                slws.HasSheetProtection = true;
                slws.SheetProtection = ProtectOptions.Clone();
            }

            return result;
        }

        /// <summary>
        /// Unprotect the currently selected worksheet.
        /// </summary>
        /// <returns>True if operation is successful. False otherwise. Note that if the worksheet is password protected or if the worksheet has no sheet protection in the first place, false is also returned.</returns>
        public bool UnprotectWorksheet()
        {
            bool result = true;
            if (slws.HasSheetProtection)
            {
                if (slws.SheetProtection.HashValue != null || slws.SheetProtection.Password != null)
                {
                    // has password protection, so return false;
                    result = false;
                }
                else
                {
                    slws.HasSheetProtection = false;
                    slws.SheetProtection = new SLSheetProtection();
                }
            }
            else
            {
                // no sheet protection, so return false (because nothing was done)
                result = false;
            }

            return result;
        }
    }
}
