using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLCacheSource
    {
        /// <summary>
        /// If true, use worksheet. If false, use consolidation. If null, use extension list.
        /// </summary>
        internal bool? IsWorksheetSource;

        // for WorksheetSource
        internal string WorksheetSourceReference { get; set; }
        internal string WorksheetSourceName { get; set; }
        internal string WorksheetSourceSheet { get; set; }
        internal string WorksheetSourceId { get; set; }

        internal SLConsolidation Consolidation { get; set; }
        internal CacheSourceExtensionList ExtensionList { get; set; }

        internal SourceValues Type { get; set; }
        internal uint ConnectionId { get; set; }

        internal SLCacheSource()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.IsWorksheetSource = true;

            this.WorksheetSourceReference = "";
            this.WorksheetSourceName = "";
            this.WorksheetSourceSheet = "";
            this.WorksheetSourceId = "";

            this.Consolidation = new SLConsolidation();
            this.ExtensionList = null;

            this.Type = SourceValues.Worksheet;
            this.ConnectionId = 0;
        }

        internal void FromCacheSource(CacheSource cs)
        {
            this.SetAllNull();

            if (cs.Type != null) this.Type = cs.Type.Value;
            if (cs.ConnectionId != null) this.ConnectionId = cs.ConnectionId.Value;

            if (cs.WorksheetSource != null)
            {
                if (cs.WorksheetSource.Reference != null) this.WorksheetSourceReference = cs.WorksheetSource.Reference.Value;
                if (cs.WorksheetSource.Name != null) this.WorksheetSourceName = cs.WorksheetSource.Name.Value;
                if (cs.WorksheetSource.Sheet != null) this.WorksheetSourceSheet = cs.WorksheetSource.Sheet.Value;
                if (cs.WorksheetSource.Id != null) this.WorksheetSourceId = cs.WorksheetSource.Id.Value;
                this.IsWorksheetSource = true;
            }
            else if (cs.Consolidation != null)
            {
                this.Consolidation.FromConsolidation(cs.Consolidation);
                this.IsWorksheetSource = false;
            }
            else if (cs.CacheSourceExtensionList != null)
            {
                this.ExtensionList = (CacheSourceExtensionList)cs.CacheSourceExtensionList.CloneNode(true);
                this.IsWorksheetSource = null;
            }
        }

        internal CacheSource ToCacheSource()
        {
            CacheSource cs = new CacheSource();

            cs.Type = this.Type;
            if (this.ConnectionId != 0) cs.ConnectionId = this.ConnectionId;

            if (this.IsWorksheetSource != null)
            {
                if (this.IsWorksheetSource.Value)
                {
                    cs.WorksheetSource = new WorksheetSource();
                    if (this.WorksheetSourceReference != null && this.WorksheetSourceReference.Length > 0) cs.WorksheetSource.Reference = this.WorksheetSourceReference;
                    if (this.WorksheetSourceName != null && this.WorksheetSourceName.Length > 0) cs.WorksheetSource.Name = this.WorksheetSourceName;
                    if (this.WorksheetSourceSheet != null && this.WorksheetSourceSheet.Length > 0) cs.WorksheetSource.Sheet = this.WorksheetSourceSheet;
                    if (this.WorksheetSourceId != null && this.WorksheetSourceId.Length > 0) cs.WorksheetSource.Id = this.WorksheetSourceId;
                }
                else
                {
                    cs.Consolidation = this.Consolidation.ToConsolidation();
                }
            }
            else
            {
                if (this.ExtensionList != null) cs.CacheSourceExtensionList = (CacheSourceExtensionList)this.ExtensionList.CloneNode(true);
            }

            return cs;
        }

        internal SLCacheSource Clone()
        {
            SLCacheSource cs = new SLCacheSource();
            cs.Type = this.Type;
            cs.ConnectionId = this.ConnectionId;

            cs.IsWorksheetSource = this.IsWorksheetSource;

            cs.WorksheetSourceReference = this.WorksheetSourceReference;
            cs.WorksheetSourceName = this.WorksheetSourceName;
            cs.WorksheetSourceSheet = this.WorksheetSourceSheet;
            cs.WorksheetSourceId = this.WorksheetSourceId;

            cs.Consolidation = this.Consolidation.Clone();

            if (this.ExtensionList != null) cs.ExtensionList = (CacheSourceExtensionList)this.ExtensionList.CloneNode(true);

            return cs;
        }
    }
}
