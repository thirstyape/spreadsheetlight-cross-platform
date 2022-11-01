using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    internal class SLPivotCacheDefinition
    {
        internal SLCacheSource CacheSource { get; set; }
        internal List<SLCacheField> CacheFields { get; set; }
        internal List<SLCacheHierarchy> CacheHierarchies { get; set; }
        internal List<SLKpi> Kpis { get; set; }

        internal bool HasTupleCache;
        internal SLTupleCache TupleCache { get; set; }

        internal List<SLCalculatedItem> CalculatedItems { get; set; }
        internal List<SLCalculatedMember> CalculatedMembers { get; set; }
        internal List<SLDimension> Dimensions { get; set; }
        internal List<SLMeasureGroup> MeasureGroups { get; set; }
        internal List<SLMeasureDimensionMap> Maps { get; set; }

        internal string Id { get; set; }
        internal bool Invalid { get; set; }
        internal bool SaveData { get; set; }
        internal bool RefreshOnLoad { get; set; }
        internal bool OptimizeMemory { get; set; }
        internal bool EnableRefresh { get; set; }
        internal string RefreshedBy { get; set; }
        internal double? RefreshedDate { get; set; }
        internal bool BackgroundQuery { get; set; }
        internal uint? MissingItemsLimit { get; set; }
        internal byte CreatedVersion { get; set; }
        internal byte RefreshedVersion { get; set; }
        internal byte MinRefreshableVersion { get; set; }
        internal uint? RecordCount { get; set; }
        internal bool UpgradeOnRefresh { get; set; }
        internal bool IsTupleCache { get; set; }
        internal bool SupportSubquery { get; set; }
        internal bool SupportAdvancedDrill { get; set; }

        internal SLPivotCacheDefinition()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.CacheSource = new SLCacheSource();
            this.CacheFields = new List<SLCacheField>();
            this.CacheHierarchies = new List<SLCacheHierarchy>();
            this.Kpis = new List<SLKpi>();
            this.HasTupleCache = false;
            this.TupleCache = new SLTupleCache();
            this.CalculatedItems = new List<SLCalculatedItem>();
            this.CalculatedMembers = new List<SLCalculatedMember>();
            this.Dimensions = new List<SLDimension>();
            this.MeasureGroups = new List<SLMeasureGroup>();
            this.Maps = new List<SLMeasureDimensionMap>();

            this.Id = "";
            this.Invalid = false;
            this.SaveData = true;
            this.RefreshOnLoad = false;
            this.OptimizeMemory = false;
            this.EnableRefresh = true;
            this.RefreshedBy = "";
            this.RefreshedDate = null;
            this.BackgroundQuery = false;
            this.MissingItemsLimit = null;

            // See SLPivotTable for similar explanation.
            this.CreatedVersion = 3;
            this.RefreshedVersion = 3;
            this.MinRefreshableVersion = 3;

            this.RecordCount = null;
            this.UpgradeOnRefresh = false;
            this.IsTupleCache = false;
            this.SupportSubquery = false;
            this.SupportAdvancedDrill = false;
        }

        internal void FromPivotCacheDefinition(PivotCacheDefinition pcd)
        {
            this.SetAllNull();

            if (pcd.Id != null) this.Id = pcd.Id.Value;
            if (pcd.Invalid != null) this.Invalid = pcd.Invalid.Value;
            if (pcd.SaveData != null) this.SaveData = pcd.SaveData.Value;
            if (pcd.RefreshOnLoad != null) this.RefreshOnLoad = pcd.RefreshOnLoad.Value;
            if (pcd.OptimizeMemory != null) this.OptimizeMemory = pcd.OptimizeMemory.Value;
            if (pcd.EnableRefresh != null) this.EnableRefresh = pcd.EnableRefresh.Value;
            if (pcd.RefreshedBy != null) this.RefreshedBy = pcd.RefreshedBy.Value;
            if (pcd.RefreshedDate != null) this.RefreshedDate = pcd.RefreshedDate.Value;
            if (pcd.BackgroundQuery != null) this.BackgroundQuery = pcd.BackgroundQuery.Value;
            if (pcd.MissingItemsLimit != null) this.MissingItemsLimit = pcd.MissingItemsLimit.Value;
            if (pcd.CreatedVersion != null) this.CreatedVersion = pcd.CreatedVersion.Value;
            if (pcd.RefreshedVersion != null) this.RefreshedVersion = pcd.RefreshedVersion.Value;
            if (pcd.MinRefreshableVersion != null) this.MinRefreshableVersion = pcd.MinRefreshableVersion.Value;
            if (pcd.RecordCount != null) this.RecordCount = pcd.RecordCount.Value;
            if (pcd.UpgradeOnRefresh != null) this.UpgradeOnRefresh = pcd.UpgradeOnRefresh.Value;
            if (pcd.IsTupleCache != null) this.IsTupleCache = pcd.IsTupleCache.Value;
            if (pcd.SupportSubquery != null) this.SupportSubquery = pcd.SupportSubquery.Value;
            if (pcd.SupportAdvancedDrill != null) this.SupportAdvancedDrill = pcd.SupportAdvancedDrill.Value;

            if (pcd.CacheSource != null) this.CacheSource.FromCacheSource(pcd.CacheSource);

            // doing one by one because it's bloody hindering awkward complicated.

            if (pcd.CacheFields != null)
            {
                SLCacheField cf;
                using (OpenXmlReader oxr = OpenXmlReader.Create(pcd.CacheFields))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(CacheField))
                        {
                            cf = new SLCacheField();
                            cf.FromCacheField((CacheField)oxr.LoadCurrentElement());
                            this.CacheFields.Add(cf);
                        }
                    }
                }
            }

            if (pcd.CacheHierarchies != null)
            {
                SLCacheHierarchy ch;
                using (OpenXmlReader oxr = OpenXmlReader.Create(pcd.CacheHierarchies))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(CacheHierarchy))
                        {
                            ch = new SLCacheHierarchy();
                            ch.FromCacheHierarchy((CacheHierarchy)oxr.LoadCurrentElement());
                            this.CacheHierarchies.Add(ch);
                        }
                    }
                }
            }

            if (pcd.Kpis != null)
            {
                SLKpi k;
                using (OpenXmlReader oxr = OpenXmlReader.Create(pcd.Kpis))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(Kpi))
                        {
                            k = new SLKpi();
                            k.FromKpi((Kpi)oxr.LoadCurrentElement());
                            this.Kpis.Add(k);
                        }
                    }
                }
            }

            if (pcd.TupleCache != null)
            {
                this.TupleCache.FromTupleCache(pcd.TupleCache);
                this.HasTupleCache = true;
            }

            if (pcd.CalculatedItems != null)
            {
                SLCalculatedItem ci;
                using (OpenXmlReader oxr = OpenXmlReader.Create(pcd.CalculatedItems))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(CalculatedItem))
                        {
                            ci = new SLCalculatedItem();
                            ci.FromCalculatedItem((CalculatedItem)oxr.LoadCurrentElement());
                            this.CalculatedItems.Add(ci);
                        }
                    }
                }
            }

            if (pcd.CalculatedMembers != null)
            {
                SLCalculatedMember cm;
                using (OpenXmlReader oxr = OpenXmlReader.Create(pcd.CalculatedMembers))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(CalculatedMember))
                        {
                            cm = new SLCalculatedMember();
                            cm.FromCalculatedMember((CalculatedMember)oxr.LoadCurrentElement());
                            this.CalculatedMembers.Add(cm);
                        }
                    }
                }
            }

            if (pcd.Dimensions != null)
            {
                SLDimension d;
                using (OpenXmlReader oxr = OpenXmlReader.Create(pcd.Dimensions))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(Dimension))
                        {
                            d = new SLDimension();
                            d.FromDimension((Dimension)oxr.LoadCurrentElement());
                            this.Dimensions.Add(d);
                        }
                    }
                }
            }

            if (pcd.MeasureGroups != null)
            {
                SLMeasureGroup mg;
                using (OpenXmlReader oxr = OpenXmlReader.Create(pcd.MeasureGroups))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(MeasureGroup))
                        {
                            mg = new SLMeasureGroup();
                            mg.FromMeasureGroup((MeasureGroup)oxr.LoadCurrentElement());
                            this.MeasureGroups.Add(mg);
                        }
                    }
                }
            }

            if (pcd.Maps != null)
            {
                SLMeasureDimensionMap mdm;
                using (OpenXmlReader oxr = OpenXmlReader.Create(pcd.Maps))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(MeasureDimensionMap))
                        {
                            mdm = new SLMeasureDimensionMap();
                            mdm.FromMeasureDimensionMap((MeasureDimensionMap)oxr.LoadCurrentElement());
                            this.Maps.Add(mdm);
                        }
                    }
                }
            }
        }

        internal PivotCacheDefinition ToPivotCacheDefinition()
        {
            PivotCacheDefinition pcd = new PivotCacheDefinition();
            if (this.Id != null && this.Id.Length > 0) pcd.Id = this.Id;
            if (this.Invalid != false) pcd.Invalid = this.Invalid;
            if (this.SaveData != true) pcd.SaveData = this.SaveData;
            if (this.RefreshOnLoad != false) pcd.RefreshOnLoad = this.RefreshOnLoad;
            if (this.OptimizeMemory != false) pcd.OptimizeMemory = this.OptimizeMemory;
            if (this.EnableRefresh != true) pcd.EnableRefresh = this.EnableRefresh;
            if (this.RefreshedBy != null && this.RefreshedBy.Length > 0) pcd.RefreshedBy = this.RefreshedBy;
            if (this.RefreshedDate != null) pcd.RefreshedDate = this.RefreshedDate.Value;
            if (this.BackgroundQuery != false) pcd.BackgroundQuery = this.BackgroundQuery;
            if (this.MissingItemsLimit != null) pcd.MissingItemsLimit = this.MissingItemsLimit.Value;
            if (this.CreatedVersion != 0) pcd.CreatedVersion = this.CreatedVersion;
            if (this.RefreshedVersion != 0) pcd.RefreshedVersion = this.RefreshedVersion;
            if (this.MinRefreshableVersion != 0) pcd.MinRefreshableVersion = this.MinRefreshableVersion;
            if (this.RecordCount != null) pcd.RecordCount = this.RecordCount.Value;
            if (this.UpgradeOnRefresh != false) pcd.UpgradeOnRefresh = this.UpgradeOnRefresh;
            if (this.IsTupleCache != false) pcd.IsTupleCache = this.IsTupleCache;
            if (this.SupportSubquery != false) pcd.SupportSubquery = this.SupportSubquery;
            if (this.SupportAdvancedDrill != false) pcd.SupportAdvancedDrill = this.SupportAdvancedDrill;

            pcd.CacheSource = this.CacheSource.ToCacheSource();

            pcd.CacheFields = new CacheFields() { Count = (uint)this.CacheFields.Count };
            foreach (SLCacheField cf in this.CacheFields)
            {
                pcd.CacheFields.Append(cf.ToCacheField());
            }

            if (this.CacheHierarchies.Count > 0)
            {
                pcd.CacheHierarchies = new CacheHierarchies() { Count = (uint)this.CacheHierarchies.Count };
                foreach (SLCacheHierarchy ch in this.CacheHierarchies)
                {
                    pcd.CacheHierarchies.Append(ch.ToCacheHierarchy());
                }
            }

            if (this.Kpis.Count > 0)
            {
                pcd.Kpis = new Kpis() { Count = (uint)this.Kpis.Count };
                foreach (SLKpi k in this.Kpis)
                {
                    pcd.Kpis.Append(k.ToKpi());
                }
            }

            if (this.HasTupleCache) pcd.TupleCache = this.TupleCache.ToTupleCache();

            if (this.CalculatedItems.Count > 0)
            {
                pcd.CalculatedItems = new CalculatedItems() { Count = (uint)this.CalculatedItems.Count };
                foreach (SLCalculatedItem ci in this.CalculatedItems)
                {
                    pcd.CalculatedItems.Append(ci.ToCalculatedItem());
                }
            }

            if (this.CalculatedMembers.Count > 0)
            {
                pcd.CalculatedMembers = new CalculatedMembers() { Count = (uint)this.CalculatedMembers.Count };
                foreach (SLCalculatedMember cm in this.CalculatedMembers)
                {
                    pcd.CalculatedMembers.Append(cm.ToCalculatedMember());
                }
            }

            if (this.Dimensions.Count > 0)
            {
                pcd.Dimensions = new Dimensions() { Count = (uint)this.Dimensions.Count };
                foreach (SLDimension d in this.Dimensions)
                {
                    pcd.Dimensions.Append(d.ToDimension());
                }
            }

            if (this.MeasureGroups.Count > 0)
            {
                pcd.MeasureGroups = new MeasureGroups() { Count = (uint)this.MeasureGroups.Count };
                foreach (SLMeasureGroup mg in this.MeasureGroups)
                {
                    pcd.MeasureGroups.Append(mg.ToMeasureGroup());
                }
            }

            if (this.Maps.Count > 0)
            {
                pcd.Maps = new Maps() { Count = (uint)this.Maps.Count };
                foreach (SLMeasureDimensionMap mdm in this.Maps)
                {
                    pcd.Maps.Append(mdm.ToMeasureDimensionMap());
                }
            }

            return pcd;
        }

        internal SLPivotCacheDefinition Clone()
        {
            SLPivotCacheDefinition pcd = new SLPivotCacheDefinition();
            pcd.Id = this.Id;
            pcd.Invalid = this.Invalid;
            pcd.SaveData = this.SaveData;
            pcd.RefreshOnLoad = this.RefreshOnLoad;
            pcd.OptimizeMemory = this.OptimizeMemory;
            pcd.EnableRefresh = this.EnableRefresh;
            pcd.RefreshedBy = this.RefreshedBy;
            pcd.RefreshedDate = this.RefreshedDate.Value;
            pcd.BackgroundQuery = this.BackgroundQuery;
            pcd.MissingItemsLimit = this.MissingItemsLimit.Value;
            pcd.CreatedVersion = this.CreatedVersion;
            pcd.RefreshedVersion = this.RefreshedVersion;
            pcd.MinRefreshableVersion = this.MinRefreshableVersion;
            pcd.RecordCount = this.RecordCount.Value;
            pcd.UpgradeOnRefresh = this.UpgradeOnRefresh;
            pcd.IsTupleCache = this.IsTupleCache;
            pcd.SupportSubquery = this.SupportSubquery;
            pcd.SupportAdvancedDrill = this.SupportAdvancedDrill;

            pcd.CacheSource = this.CacheSource.Clone();

            pcd.CacheFields = new List<SLCacheField>();
            foreach (SLCacheField cf in this.CacheFields)
            {
                pcd.CacheFields.Add(cf.Clone());
            }

            pcd.CacheHierarchies = new List<SLCacheHierarchy>();
            foreach (SLCacheHierarchy ch in this.CacheHierarchies)
            {
                pcd.CacheHierarchies.Add(ch.Clone());
            }

            pcd.Kpis = new List<SLKpi>();
            foreach (SLKpi k in this.Kpis)
            {
                pcd.Kpis.Add(k.Clone());
            }

            pcd.HasTupleCache = this.HasTupleCache;
            pcd.TupleCache = this.TupleCache.Clone();

            pcd.CalculatedItems = new List<SLCalculatedItem>();
            foreach (SLCalculatedItem ci in this.CalculatedItems)
            {
                pcd.CalculatedItems.Add(ci.Clone());
            }

            pcd.CalculatedMembers = new List<SLCalculatedMember>();
            foreach (SLCalculatedMember cm in this.CalculatedMembers)
            {
                pcd.CalculatedMembers.Add(cm.Clone());
            }

            pcd.Dimensions = new List<SLDimension>();
            foreach (SLDimension d in this.Dimensions)
            {
                pcd.Dimensions.Add(d.Clone());
            }

            pcd.MeasureGroups = new List<SLMeasureGroup>();
            foreach (SLMeasureGroup mg in this.MeasureGroups)
            {
                pcd.MeasureGroups.Add(mg.Clone());
            }

            pcd.Maps = new List<SLMeasureDimensionMap>();
            foreach (SLMeasureDimensionMap mdm in this.Maps)
            {
                pcd.Maps.Add(mdm.Clone());
            }

            return pcd;
        }
    }
}
