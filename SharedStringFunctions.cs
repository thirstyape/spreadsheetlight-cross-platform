using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// Get existing shared strings. WARNING: This is only a snapshot. Any changes made to the returned result are not used.
        /// </summary>
        /// <returns>A list of existing shared strings.</returns>
        public List<SLRstType> GetSharedStrings()
        {
            List<SLRstType> result = new List<SLRstType>();
            SLRstType rst = new SLRstType();
            for (int i = 0; i < listSharedString.Count; ++i)
            {
                rst.FromHash(listSharedString[i]);
                result.Add(rst.Clone());
            }

            return result;
        }

        /// <summary>
        /// Get existing shared strings in a list of SharedStringItem objects. WARNING: This is only a snapshot. Any changes made to the returned result are not used.
        /// </summary>
        /// <returns>A list of existing SharedStringItem objects.</returns>
        public List<SharedStringItem> GetSharedStringItems()
        {
            List<SharedStringItem> result = new List<SharedStringItem>();
            SharedStringItem ssi;

            for (int i = 0; i < listSharedString.Count; ++i)
            {
                ssi = new SharedStringItem();
                ssi.InnerXml = listSharedString[i];
                result.Add(ssi);
            }

            return result;
        }

        internal void LoadSharedStringTable()
        {
            countSharedString = 0;
            listSharedString = new List<string>();
            dictSharedStringHash = new Dictionary<string, int>();
            hsUniqueSharedString = new HashSet<string>();

            if (wbp.SharedStringTablePart != null)
            {
                OpenXmlReader oxr = OpenXmlReader.Create(wbp.SharedStringTablePart);
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(SharedStringItem))
                    {
                        this.ForceSaveToSharedStringTable((SharedStringItem)oxr.LoadCurrentElement());
                    }
                }
                oxr.Dispose();

                countSharedString = listSharedString.Count;
            }
        }

        internal void WriteSharedStringTable()
        {
            if (wbp.SharedStringTablePart != null)
            {
                if (listSharedString.Count > countSharedString)
                {
                    if (gbWriteUniqueSharedStringCount)
                    {
                        wbp.SharedStringTablePart.SharedStringTable.Count = (uint)listSharedString.Count;
                        wbp.SharedStringTablePart.SharedStringTable.UniqueCount = (uint)hsUniqueSharedString.Count;
                    }
                    else
                    {
                        wbp.SharedStringTablePart.SharedStringTable.Count = null;
                        wbp.SharedStringTablePart.SharedStringTable.UniqueCount = null;
                    }

                    int diff = listSharedString.Count - countSharedString;
                    for (int i = 0; i < diff; ++i)
                    {
                        wbp.SharedStringTablePart.SharedStringTable.Append(new SharedStringItem()
                        {
                            InnerXml = listSharedString[i + countSharedString]
                        });
                    }

                    wbp.SharedStringTablePart.SharedStringTable.Save();
                }
            }
            else
            {
                if (listSharedString.Count > 0)
                {
                    SharedStringTablePart sstp = wbp.AddNewPart<SharedStringTablePart>();
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (StreamWriter sw = new StreamWriter(ms))
                        {
                            if (gbWriteUniqueSharedStringCount)
                            {
                                sw.Write("<x:sst count=\"{0}\" uniqueCount=\"{1}\" xmlns:x=\"{2}\">", listSharedString.Count, hsUniqueSharedString.Count, SLConstants.NamespaceX);
                            }
                            else
                            {
                                sw.Write("<x:sst xmlns:x=\"{0}\">", SLConstants.NamespaceX);
                            }
                            
                            for (int i = 0; i < listSharedString.Count; ++i)
                            {
                                sw.Write("<x:si>{0}</x:si>", listSharedString[i]);
                            }
                            sw.Write("</x:sst>");
                            sw.Flush();
                            ms.Position = 0;
                            sstp.FeedData(ms);
                        }
                    }
                }
            }
        }
    }
}
