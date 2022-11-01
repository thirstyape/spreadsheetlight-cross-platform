using System;
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLight.Charts
{
    internal class SLMultiLevelStringReference
    {
        internal string Formula { get; set; }
        internal SLMultiLevelStringCache MultiLevelStringCache { get; set; }

        internal SLMultiLevelStringReference()
        {
            this.Formula = string.Empty;
            this.MultiLevelStringCache = new SLMultiLevelStringCache();
        }

        internal C.MultiLevelStringReference ToMultiLevelStringReference()
        {
            C.MultiLevelStringReference mlsr = new C.MultiLevelStringReference();
            mlsr.Formula = new C.Formula(this.Formula);
            mlsr.MultiLevelStringCache = this.MultiLevelStringCache.ToMultiLevelStringCache();

            return mlsr;
        }

        internal SLMultiLevelStringReference Clone()
        {
            SLMultiLevelStringReference mlsr = new SLMultiLevelStringReference();
            mlsr.Formula = this.Formula;
            mlsr.MultiLevelStringCache = this.MultiLevelStringCache.Clone();

            return mlsr;
        }
    }

    internal class SLMultiLevelStringCache
    {
        internal uint PointCount { get; set; }

        internal List<SLLevel> Levels { get; set; }

        internal SLMultiLevelStringCache()
        {
            this.PointCount = 0;
            this.Levels = new List<SLLevel>();
        }

        internal C.MultiLevelStringCache ToMultiLevelStringCache()
        {
            C.MultiLevelStringCache mlsc = new C.MultiLevelStringCache();
            mlsc.PointCount = new C.PointCount() { Val = this.PointCount };

            C.Level lvl;
            int i, j;
            for (i = 0; i < this.Levels.Count; ++i)
            {
                lvl = new C.Level();
                for (j = 0; j < this.Levels[i].Points.Count; ++j)
                {
                    lvl.Append(this.Levels[i].Points[j].ToStringPoint());
                }
                mlsc.Append(lvl);
            }

            return mlsc;
        }

        internal SLMultiLevelStringCache Clone()
        {
            SLMultiLevelStringCache mlsc = new SLMultiLevelStringCache();
            mlsc.PointCount = this.PointCount;
            for (int i = 0; i < this.Levels.Count; ++i)
            {
                mlsc.Levels.Add(this.Levels[i].Clone());
            }

            return mlsc;
        }
    }

    internal class SLLevel
    {
        internal List<SLStringPoint> Points { get; set; }

        internal SLLevel()
        {
            this.Points = new List<SLStringPoint>();
        }

        internal SLLevel Clone()
        {
            SLLevel lvl = new SLLevel();
            for (int i = 0; i < this.Points.Count; ++i)
            {
                lvl.Points.Add(this.Points[i].Clone());
            }

            return lvl;
        }
    }
}
