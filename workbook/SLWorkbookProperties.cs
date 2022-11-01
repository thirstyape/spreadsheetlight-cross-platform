using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLWorkbookProperties
    {
        internal bool HasWorkbookProperties
        {
            get
            {
                return this.bDate1904 != null || this.bDateCompatibility != null || this.vShowObjects != null
                    || this.bShowBorderUnselectedTables != null || this.bFilterPrivacy != null || this.bPromptedSolutions != null
                    || this.bShowInkAnnotation != null || this.bBackupFile != null || this.bSaveExternalLinkValues != null
                    || this.vUpdateLinks != null || this.sCodeName != null || this.bHidePivotFieldList != null
                    || this.bShowPivotChartFilter != null || this.bAllowRefreshQuery != null || this.bPublishItems != null
                    || this.bCheckCompatibility != null || this.bAutoCompressPictures != null || this.bRefreshAllConnections != null
                    || this.iDefaultThemeVersion != null;
            }
        }

        private bool? bDate1904;
        internal bool Date1904
        {
            get { return bDate1904 ?? false; }
            set { bDate1904 = value; }
        }

        private bool? bDateCompatibility;
        internal bool DateCompatibility
        {
            get { return bDateCompatibility ?? true; }
            set { bDateCompatibility = value; }
        }

        private ObjectDisplayValues? vShowObjects;
        internal ObjectDisplayValues ShowObjects
        {
            get { return vShowObjects ?? ObjectDisplayValues.All; }
            set { vShowObjects = value; }
        }

        private bool? bShowBorderUnselectedTables;
        internal bool ShowBorderUnselectedTables
        {
            get { return bShowBorderUnselectedTables ?? true; }
            set { bShowBorderUnselectedTables = value; }
        }

        private bool? bFilterPrivacy;
        internal bool FilterPrivacy
        {
            get { return bFilterPrivacy ?? false; }
            set { bFilterPrivacy = value; }
        }

        private bool? bPromptedSolutions;
        internal bool PromptedSolutions
        {
            get { return bPromptedSolutions ?? false; }
            set { bPromptedSolutions = value; }
        }

        private bool? bShowInkAnnotation;
        internal bool ShowInkAnnotation
        {
            get { return bShowInkAnnotation ?? true; }
            set { bShowInkAnnotation = value; }
        }

        private bool? bBackupFile;
        internal bool BackupFile
        {
            get { return bBackupFile ?? false; }
            set { bBackupFile = value; }
        }

        private bool? bSaveExternalLinkValues;
        internal bool SaveExternalLinkValues
        {
            get { return bSaveExternalLinkValues ?? true; }
            set { bSaveExternalLinkValues = value; }
        }

        private UpdateLinksBehaviorValues? vUpdateLinks;
        internal UpdateLinksBehaviorValues UpdateLinks
        {
            get { return vUpdateLinks ?? UpdateLinksBehaviorValues.UserSet; }
            set { vUpdateLinks = value; }
        }

        private string sCodeName;
        internal string CodeName
        {
            get { return sCodeName ?? ""; }
            set { sCodeName = value; }
        }

        private bool? bHidePivotFieldList;
        internal bool HidePivotFieldList
        {
            get { return bHidePivotFieldList ?? false; }
            set { bHidePivotFieldList = value; }
        }

        private bool? bShowPivotChartFilter;
        internal bool ShowPivotChartFilter
        {
            get { return bShowPivotChartFilter ?? false; }
            set { bShowPivotChartFilter = value; }
        }

        private bool? bAllowRefreshQuery;
        internal bool AllowRefreshQuery
        {
            get { return bAllowRefreshQuery ?? false; }
            set { bAllowRefreshQuery = value; }
        }

        private bool? bPublishItems;
        internal bool PublishItems
        {
            get { return bPublishItems ?? false; }
            set { bPublishItems = value; }
        }

        private bool? bCheckCompatibility;
        internal bool CheckCompatibility
        {
            get { return bCheckCompatibility ?? false; }
            set { bCheckCompatibility = value; }
        }

        private bool? bAutoCompressPictures;
        internal bool AutoCompressPictures
        {
            get { return bAutoCompressPictures ?? true; }
            set { bAutoCompressPictures = value; }
        }

        private bool? bRefreshAllConnections;
        internal bool RefreshAllConnections
        {
            get { return bRefreshAllConnections ?? false; }
            set { bRefreshAllConnections = value; }
        }

        private uint? iDefaultThemeVersion;
        internal uint DefaultThemeVersion
        {
            get { return iDefaultThemeVersion ?? 0; }
            set { iDefaultThemeVersion = value; }
        }

        internal SLWorkbookProperties()
        {
            this.SetAllNull();
        }

        internal void SetAllNull()
        {
            this.bDate1904 = null;
            this.bDateCompatibility = null;
            this.vShowObjects = null;
            this.bShowBorderUnselectedTables = null;
            this.bFilterPrivacy = null;
            this.bPromptedSolutions = null;
            this.bShowInkAnnotation = null;
            this.bBackupFile = null;
            this.bSaveExternalLinkValues = null;
            this.vUpdateLinks = null;
            this.sCodeName = null;
            this.bHidePivotFieldList = null;
            this.bShowPivotChartFilter = null;
            this.bAllowRefreshQuery = null;
            this.bPublishItems = null;
            this.bCheckCompatibility = null;
            this.bAutoCompressPictures = null;
            this.bRefreshAllConnections = null;
            this.iDefaultThemeVersion = null;
        }

        internal void FromWorkbookProperties(WorkbookProperties wp)
        {
            this.SetAllNull();
            if (wp.Date1904 != null) this.Date1904 = wp.Date1904.Value;
            if (wp.DateCompatibility != null) this.DateCompatibility = wp.DateCompatibility.Value;
            if (wp.ShowObjects != null) this.ShowObjects = wp.ShowObjects.Value;
            if (wp.ShowBorderUnselectedTables != null) this.ShowBorderUnselectedTables = wp.ShowBorderUnselectedTables.Value;
            if (wp.FilterPrivacy != null) this.FilterPrivacy = wp.FilterPrivacy.Value;
            if (wp.PromptedSolutions != null) this.PromptedSolutions = wp.PromptedSolutions.Value;
            if (wp.ShowInkAnnotation != null) this.ShowInkAnnotation = wp.ShowInkAnnotation.Value;
            if (wp.BackupFile != null) this.BackupFile = wp.BackupFile.Value;
            if (wp.SaveExternalLinkValues != null) this.SaveExternalLinkValues = wp.SaveExternalLinkValues.Value;
            if (wp.UpdateLinks != null) this.UpdateLinks = wp.UpdateLinks.Value;
            if (wp.CodeName != null) this.CodeName = wp.CodeName.Value;
            if (wp.HidePivotFieldList != null) this.HidePivotFieldList = wp.HidePivotFieldList.Value;
            if (wp.ShowPivotChartFilter != null) this.ShowPivotChartFilter = wp.ShowPivotChartFilter.Value;
            if (wp.AllowRefreshQuery != null) this.AllowRefreshQuery = wp.AllowRefreshQuery.Value;
            if (wp.PublishItems != null) this.PublishItems = wp.PublishItems.Value;
            if (wp.CheckCompatibility != null) this.CheckCompatibility = wp.CheckCompatibility.Value;
            if (wp.AutoCompressPictures != null) this.AutoCompressPictures = wp.AutoCompressPictures.Value;
            if (wp.RefreshAllConnections != null) this.RefreshAllConnections = wp.RefreshAllConnections.Value;
            if (wp.DefaultThemeVersion != null) this.DefaultThemeVersion = wp.DefaultThemeVersion.Value;
        }

        internal WorkbookProperties ToWorkbookProperties()
        {
            WorkbookProperties wp = new WorkbookProperties();
            if (this.bDate1904 != null) wp.Date1904 = this.bDate1904.Value;
            if (this.bDateCompatibility != null) wp.DateCompatibility = this.bDateCompatibility.Value;
            if (this.vShowObjects != null) wp.ShowObjects = this.vShowObjects.Value;
            if (this.bShowBorderUnselectedTables != null) wp.ShowBorderUnselectedTables = this.bShowBorderUnselectedTables.Value;
            if (this.bFilterPrivacy != null) wp.FilterPrivacy = this.bFilterPrivacy.Value;
            if (this.bPromptedSolutions != null) wp.PromptedSolutions = this.bPromptedSolutions.Value;
            if (this.bShowInkAnnotation != null) wp.ShowInkAnnotation = this.bShowInkAnnotation.Value;
            if (this.bBackupFile != null) wp.BackupFile = this.bBackupFile.Value;
            if (this.bSaveExternalLinkValues != null) wp.SaveExternalLinkValues = this.bSaveExternalLinkValues.Value;
            if (this.vUpdateLinks != null) wp.UpdateLinks = this.vUpdateLinks.Value;
            if (this.sCodeName != null) wp.CodeName = this.sCodeName;
            if (this.bHidePivotFieldList != null) wp.HidePivotFieldList = this.bHidePivotFieldList.Value;
            if (this.bShowPivotChartFilter != null) wp.ShowPivotChartFilter = this.bShowPivotChartFilter.Value;
            if (this.bAllowRefreshQuery != null) wp.AllowRefreshQuery = this.bAllowRefreshQuery.Value;
            if (this.bPublishItems != null) wp.PublishItems = this.bPublishItems.Value;
            if (this.bCheckCompatibility != null) wp.CheckCompatibility = this.bCheckCompatibility.Value;
            if (this.bAutoCompressPictures != null) wp.AutoCompressPictures = this.bAutoCompressPictures.Value;
            if (this.bRefreshAllConnections != null) wp.RefreshAllConnections = this.bRefreshAllConnections.Value;
            if (this.iDefaultThemeVersion != null) wp.DefaultThemeVersion = this.iDefaultThemeVersion.Value;

            return wp;
        }
    }
}
